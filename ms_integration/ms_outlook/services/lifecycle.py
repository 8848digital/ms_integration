# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""
Account lifecycle after creation: detail updates, offboarding (licenses,
groups, block sign-in, sign out), rehire, password reset and deletion of
long-disabled accounts.
"""

import frappe
from frappe import _
from frappe.utils import add_days, now_datetime

from ms_integration.ms_core.services.graph import GraphClient, friendly
from ms_integration.ms_core.services.logs import log
from ms_integration.ms_outlook.services import account_api
from ms_integration.ms_outlook.services import directory_api as api
from ms_integration.ms_outlook.services.access import grant, rule_access
from ms_integration.ms_outlook.services.directory import refresh_account
from ms_integration.ms_outlook.services.payload import manager_object_id, profile
from ms_integration.ms_outlook.services.settings import get_settings
from ms_integration.ms_outlook.services.welcome import new_password, send_welcome

ACCOUNT = "Microsoft 365 Account"


def update_account(account_name: str) -> None:
	"""
	Send the employee's current details and manager to Microsoft 365.

	Parameters:
	        account_name (str, required): Microsoft 365 Account name.

	Returns:
	        None
	"""
	account = frappe.get_doc(ACCOUNT, account_name)
	if account.status != "Active":
		return
	settings = get_settings()
	employee = frappe.get_doc("Employee", account.employee)
	client = GraphClient()
	try:
		api.update_user(client, account.object_id, profile(employee, settings))
		manager = manager_object_id(employee, settings)
		if manager:
			api.set_manager(client, account.object_id, manager)
		log("Update Account", account=account.name, details=_("Details and manager updated"))
	except Exception as e:
		log("Update Account", False, account.name, error=e)
		account.db_set("error", friendly(e))


def disable_account(account_name: str) -> None:
	"""
	Offboarding: remove licenses and groups (as configured), block sign-in and
	sign out everywhere. The account is Disabled only when sign-in is blocked.

	Parameters:
	        account_name (str, required): Microsoft 365 Account name.

	Returns:
	        None
	"""
	account = frappe.get_doc(ACCOUNT, account_name)
	if account.status != "Active":
		return
	settings, client = get_settings(), GraphClient()
	user_id = account.object_id
	errors = []
	if settings.remove_licenses_on_exit:
		attempt(
			errors,
			lambda: api.assign_licenses(client, user_id, remove=api.user_license_skus(client, user_id)),
		)
	if settings.remove_groups_on_exit:
		attempt(errors, lambda: remove_groups(client, user_id))
	blocked = attempt(errors, lambda: account_api.set_sign_in(client, user_id, False))
	if blocked and settings.revoke_sessions:
		attempt(errors, lambda: account_api.revoke_sessions(client, user_id))
	if blocked:
		account.db_set({"status": "Disabled", "account_enabled": 0, "disabled_on": now_datetime()})
	account.db_set("error", "\n".join(errors) or None)
	log("Disable Account", blocked and not errors, account.name, error="\n".join(errors) or None)
	safe_refresh(client, account)


def attempt(errors: list, step) -> bool:
	"""
	Run one offboarding step, collecting its error instead of stopping.

	Parameters:
	        errors (list, required): Collected error messages.
	        step (Callable, required): Step to run.

	Returns:
	        bool: True when the step succeeded.
	"""
	try:
		step()
		return True
	except Exception as e:
		errors.append(friendly(e))
		return False


def remove_groups(client: GraphClient, user_id: str) -> None:
	"""
	Remove the user from every group whose members can be changed.

	Parameters:
	        client (GraphClient, required): Graph client.
	        user_id (str, required): User object id.

	Returns:
	        None
	"""
	failed = []
	for group in api.user_groups(client, user_id):
		dynamic = "DynamicMembership" in (group.get("groupTypes") or [])
		mail_only = group.get("mailEnabled") and "Unified" not in (group.get("groupTypes") or [])
		if dynamic or mail_only:
			continue
		try:
			api.remove_from_group(client, group["id"], user_id)
		except Exception as e:
			failed.append(f"{group.get('displayName')}: {friendly(e)}")
	if failed:
		frappe.throw(_("Could not remove from groups: {0}").format("; ".join(failed)))


def enable_account(account_name: str) -> None:
	"""
	Rehire: allow sign-in, give the rule licenses and groups, new password + welcome email.

	Parameters:
	        account_name (str, required): Microsoft 365 Account name.

	Returns:
	        None
	"""
	account = frappe.get_doc(ACCOUNT, account_name)
	if account.status != "Disabled":
		return
	client = GraphClient()
	try:
		account_api.set_sign_in(client, account.object_id, True)
	except Exception as e:
		log("Enable Account", False, account.name, error=e)
		account.db_set("error", friendly(e))
		return
	account.db_set({"status": "Active", "account_enabled": 1, "disabled_on": None, "error": None})
	log("Enable Account", account=account.name)
	employee = frappe.get_doc("Employee", account.employee)
	errors = grant(client, account, *rule_access(employee))
	if get_settings().send_welcome_email:
		errors += reset_password(account.name, client)
	safe_refresh(client, account)
	if errors:
		account.db_set("error", "\n".join(errors))


def reset_password(account_name: str, client: GraphClient | None = None) -> list[str]:
	"""
	New temporary password (change at next sign-in) emailed to the employee.

	Parameters:
	        account_name (str, required): Microsoft 365 Account name.
	        client (GraphClient, optional): Graph client.

	Returns:
	        list[str]: Error messages.
	"""
	account = frappe.get_doc(ACCOUNT, account_name)
	client = client or GraphClient()
	settings = get_settings()
	password = new_password()
	try:
		account_api.set_password(client, account.object_id, password)
		employee = frappe.get_doc("Employee", account.employee)
		recipient = send_welcome(client, employee, account.user_principal_name, password, settings)
	except Exception as e:
		log("Reset Password", False, account.name, error=e)
		return [friendly(e)]
	log(
		"Reset Password", account=account.name, details=_("New password sent to {0}").format(recipient)
	)
	return []


def delete_expired() -> int:
	"""
	Delete accounts disabled longer than the configured number of days.

	Returns:
	        int: Accounts deleted.
	"""
	days = get_settings().delete_after_days
	if not days:
		return 0
	cutoff = add_days(now_datetime(), -days)
	names = frappe.get_all(
		ACCOUNT, {"status": "Disabled", "disabled_on": ("<", cutoff)}, pluck="name"
	)
	client = GraphClient()
	for name in names:
		account = frappe.get_doc(ACCOUNT, name)
		try:
			account_api.delete_user(client, account.object_id)
		except Exception as e:
			log("Delete Account", False, name, error=e)
			continue
		account.db_set({"status": "Deleted", "account_enabled": 0})
		log("Delete Account", account=name, details=account.user_principal_name)
	return len(names)


def safe_refresh(client: GraphClient, account) -> None:
	"""
	Refresh rows from Microsoft 365; a failure is only logged.

	Parameters:
	        client (GraphClient, required): Graph client.
	        account (Document, required): Microsoft 365 Account.

	Returns:
	        None
	"""
	try:
		account.reload()
		refresh_account(client, account)
	except Exception as e:
		log("Directory Sync", False, account.name, error=e)

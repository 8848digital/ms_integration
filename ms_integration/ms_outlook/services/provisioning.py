# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""
Onboarding: create the employee's Microsoft 365 account.

Free address (ERPNext + Microsoft 365) → user with a random temporary password
(change at first sign-in) → manager → email saved on the Employee (and an
ERPNext User if configured) → rule licenses and groups → welcome email.
"""

import frappe
from frappe import _
from frappe.utils import now_datetime

from ms_integration.ms_core.services.graph import GraphClient, GraphError, friendly
from ms_integration.ms_core.services.logs import log
from ms_integration.ms_core.services.settings import lines
from ms_integration.ms_outlook.services import directory_api as api
from ms_integration.ms_outlook.services.access import grant, rule_access
from ms_integration.ms_outlook.services.addresses import choose_address, is_free
from ms_integration.ms_outlook.services.directory import refresh_account
from ms_integration.ms_outlook.services.payload import profile
from ms_integration.ms_outlook.services.provisioning_steps import link_email, set_manager, welcome
from ms_integration.ms_outlook.services.settings import get_settings, is_enabled
from ms_integration.ms_outlook.services.welcome import new_password

ACCOUNT = "Microsoft 365 Account"
TAKEN = "already exists"
MAX_TRIES = 5


def account_for(employee: str):
	"""
	The employee's Microsoft 365 Account, created as Pending when missing.

	Parameters:
	        employee (str, required): Employee name.

	Returns:
	        Document: Microsoft 365 Account.
	"""
	if frappe.db.exists(ACCOUNT, employee):
		return frappe.get_doc(ACCOUNT, employee)
	account = frappe.get_doc({"doctype": ACCOUNT, "employee": employee, "status": "Pending"})
	account.flags.from_sync = True
	return account.insert(ignore_permissions=True)


def provision(employee_name: str) -> None:
	"""
	Create the Microsoft 365 account (background job). Failures leave the
	account Failed with the reason; nothing is retried silently.

	Parameters:
	        employee_name (str, required): Employee name.

	Returns:
	        None
	"""
	settings = get_settings()
	account = account_for(employee_name)
	if not is_enabled() or account.status in ("Active", "Disabled"):
		return
	employee = frappe.get_doc("Employee", employee_name)
	client = GraphClient()
	try:
		check_required_fields(employee, settings)
		password = create_user(client, employee, account, settings)
	except Exception as e:
		mark_failed(account, e)
		return
	finish(client, employee, account, settings, password)


def check_required_fields(employee, settings) -> None:
	"""
	Stop when a required Employee field is empty.

	Parameters:
	        employee (Document, required): Employee.
	        settings (Document, required): Microsoft 365 Settings.

	Returns:
	        None
	"""
	missing = [
		employee.meta.get_label(f) for f in lines(settings.required_fields) if not employee.get(f)
	]
	if missing:
		frappe.throw(_("Fill these Employee fields first: {0}").format(", ".join(missing)))


def create_user(client: GraphClient, employee, account, settings) -> str:
	"""
	Create the user with the first free address and save it on the account.

	Parameters:
	        client (GraphClient, required): Graph client.
	        employee (Document, required): Employee.
	        account (Document, required): Microsoft 365 Account.
	        settings (Document, required): Microsoft 365 Settings.

	Returns:
	        str: The temporary password (only kept in memory for the welcome email).
	"""
	requested = (account.user_principal_name or "").strip().lower()
	if requested and not is_free(client, requested, employee.name, settings):
		frappe.throw(_("{0} is already used in ERPNext or Microsoft 365.").format(requested))
	account.db_set("status", "Creating")
	skip = []
	for _try in range(MAX_TRIES):
		email = requested or choose_address(client, employee, settings, skip)
		password = new_password()
		body = {
			**{k: v for k, v in profile(employee, settings).items() if v not in (None, "", [])},
			"accountEnabled": True,
			"userPrincipalName": email,
			"mailNickname": email.split("@")[0],
			"passwordProfile": {"forceChangePasswordNextSignIn": True, "password": password},
		}
		try:
			user = api.create_user(client, body)
		except GraphError as e:
			if requested or TAKEN not in e.message:
				raise
			skip.append(email)
			continue
		save_created(account, user, email)
		return password
	frappe.throw(_("Microsoft 365 kept reporting the address as taken; try again later."))


def save_created(account, user: dict, email: str) -> None:
	"""
	Record the new user and commit at once, so a later failure can't lose it.

	Parameters:
	        account (Document, required): Microsoft 365 Account.
	        user (dict, required): Created Graph user.
	        email (str, required): Sign-in address.

	Returns:
	        None
	"""
	account.update(
		{
			"object_id": user["id"],
			"user_principal_name": email,
			"status": "Active",
			"account_enabled": 1,
			"created_on": now_datetime(),
			"error": None,
		}
	)
	account.flags.from_sync = True
	account.save(ignore_permissions=True)
	log("Create Account", account=account.name, details=email)
	# The Microsoft 365 user now exists: keep its record even if a later step fails.
	frappe.db.commit()  # nosemgrep


def finish(client: GraphClient, employee, account, settings, password: str) -> None:
	"""
	Steps after creation; each failure is logged and shown on the account.

	Parameters:
	        client (GraphClient, required): Graph client.
	        employee (Document, required): Employee.
	        account (Document, required): Microsoft 365 Account (Active).
	        settings (Document, required): Microsoft 365 Settings.
	        password (str, required): Temporary password.

	Returns:
	        None
	"""
	errors = []
	steps = [
		("Update Account", lambda: set_manager(client, employee, account, settings)),
		("Update Account", lambda: link_email(employee, account, settings)),
		("Welcome Email", lambda: welcome(client, employee, account, settings, password)),
	]
	licenses, groups = rule_access(employee)
	extra_l = [r.license for r in account.licenses if r.license not in licenses]
	extra_g = [r.group for r in account.groups if r.group not in groups]
	errors += grant(client, account, licenses + extra_l, groups + extra_g)
	for action, step in steps:
		try:
			step()
		except Exception as e:
			errors.append(friendly(e))
			log(action, False, account.name, error=e)
	refresh_account(client, account)
	if errors:
		account.db_set("error", "\n".join(errors))


def mark_failed(account, error: Exception) -> None:
	"""
	Mark the account Failed with the reason.

	Parameters:
	        account (Document, required): Microsoft 365 Account.
	        error (Exception, required): What went wrong.

	Returns:
	        None
	"""
	account.reload()
	if account.status == "Active":
		return
	account.db_set({"status": "Failed", "error": friendly(error)})
	log("Create Account", False, account.name, account.user_principal_name, error)

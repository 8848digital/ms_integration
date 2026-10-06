# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Microsoft 365 Account actions for the Employee and Account forms."""

import frappe
from frappe import _

from ms_integration.ms_core.services.graph import GraphClient
from ms_integration.ms_outlook.services import lifecycle
from ms_integration.ms_outlook.services.directory import refresh_account
from ms_integration.ms_outlook.services.provisioning import account_for
from ms_integration.ms_outlook.services.settings import is_enabled
from ms_integration.ms_outlook.tasks import enqueue
from ms_integration.utils.api_handlers.response_formatter import api_response

ACCOUNT = "Microsoft 365 Account"


@frappe.whitelist(methods=["GET"])
def get_status(employee: str):
	"""
	        The employee's Microsoft 365 account status, for the Employee form.
	        Empty data when the user may not read accounts or none exists.

	        **Endpoint:** `/api/method/ms_integration.ms_outlook.api.v1.account.get_status`
	        **HTTP Method:** GET
	        **Parameters:**
	                - employee (str, required): Employee name
	        **Response:**
	```json
	                {"status": true, "status_code": 200, "message": "Account found",
	                 "data": {"enabled": true, "account": "HR-EMP-00001", "status": "Active",
	                          "email": "john.s@example.com", "error": null}, "errors": null}
	```
	"""
	frappe.has_permission("Employee", "read", employee, throw=True)
	data = {"enabled": is_enabled()}
	if frappe.has_permission(ACCOUNT, "read") and frappe.db.exists(ACCOUNT, employee):
		values = frappe.db.get_value(
			ACCOUNT, employee, ["status", "user_principal_name", "error"], as_dict=1
		)
		data.update(
			account=employee, status=values.status, email=values.user_principal_name, error=values.error
		)
	return api_response(
		message=_("Account found") if data.get("account") else _("No account"), data=data
	)


@frappe.whitelist(methods=["POST"])
def create_account(employee: str):
	"""
	        Create (or retry creating) the employee's Microsoft 365 account in the background.

	        **Endpoint:** `/api/method/ms_integration.ms_outlook.api.v1.account.create_account`
	        **HTTP Method:** POST
	        **Parameters:**
	                - employee (str, required): Employee name
	        **Response:**
	```json
	                {"status": true, "status_code": 200, "message": "Creating the Microsoft 365 account…",
	                 "data": {"account": "HR-EMP-00001"}, "errors": null}
	```
	"""
	frappe.has_permission(ACCOUNT, "create", throw=True)
	frappe.has_permission("Employee", "read", employee, throw=True)
	check_enabled()
	account = account_for(employee)
	if account.status in ("Active", "Disabled", "Creating"):
		frappe.throw(_("The account is already {0}.").format(_(account.status)))
	enqueue("provision_account", employee, employee=employee)
	return api_response(
		message=_("Creating the Microsoft 365 account…"), data={"account": account.name}
	)


@frappe.whitelist(methods=["POST"])
def run_action(account: str, action: str):
	"""
	        Run an account action: disable (offboard), enable (rehire), reset_password
	        (new temporary password emailed) or refresh (re-read from Microsoft 365).

	        **Endpoint:** `/api/method/ms_integration.ms_outlook.api.v1.account.run_action`
	        **HTTP Method:** POST
	        **Parameters:**
	                - account (str, required): Microsoft 365 Account name
	                - action (str, required): disable | enable | reset_password | refresh
	        **Response:**
	```json
	                {"status": true, "status_code": 200, "message": "Password reset and emailed.",
	                 "data": {"account": "HR-EMP-00001"}, "errors": null}
	```
	"""
	frappe.has_permission(ACCOUNT, "write", account, throw=True)
	check_enabled()
	doc = frappe.get_doc(ACCOUNT, account)
	if action in ("disable", "enable"):
		needed = "Active" if action == "disable" else "Disabled"
		if doc.status != needed:
			frappe.throw(_("Only an {0} account can be changed this way.").format(_(needed)))
		enqueue(f"{action}_account", account, account=account)
		message = _("Disabling the account…") if action == "disable" else _("Enabling the account…")
	elif action == "reset_password":
		if doc.status != "Active":
			frappe.throw(_("Only an Active account can get a new password."))
		errors = lifecycle.reset_password(account)
		if errors:
			frappe.throw("<br>".join(errors))
		message = _("Password reset and emailed.")
	elif action == "refresh":
		if not doc.object_id:
			frappe.throw(_("The account doesn't exist in Microsoft 365 yet."))
		refresh_account(GraphClient(), doc)
		message = _("Refreshed from Microsoft 365.")
	else:
		frappe.throw(_("Unknown action {0}.").format(action))
	return api_response(message=message, data={"account": account})


def check_enabled() -> None:
	"""
	Stop when the integration is switched off.

	Returns:
	        None
	"""
	if not is_enabled():
		frappe.throw(_("Microsoft 365 integration is not enabled."))

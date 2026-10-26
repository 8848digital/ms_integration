# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Steps after a new account exists: manager, email on the Employee, welcome email."""

import frappe
from frappe import _

from ms_integration.ms_core.services.graph import GraphClient
from ms_integration.ms_core.services.logs import log
from ms_integration.ms_outlook.services import directory_api as api
from ms_integration.ms_outlook.services.payload import manager_object_id
from ms_integration.ms_outlook.services.welcome import send_welcome


def set_manager(client: GraphClient, employee, account, settings) -> None:
	"""
	Set the manager when the manager already has an account.

	Parameters:
	        client (GraphClient, required): Graph client.
	        employee (Document, required): Employee.
	        account (Document, required): Microsoft 365 Account.
	        settings (Document, required): Microsoft 365 Settings.

	Returns:
	        None
	"""
	manager = manager_object_id(employee, settings)
	if manager:
		api.set_manager(client, account.object_id, manager)


def link_email(employee, account, settings) -> None:
	"""
	Save the address on the Employee and, if configured, create/link an ERPNext User.

	Parameters:
	        employee (Document, required): Employee.
	        account (Document, required): Microsoft 365 Account.
	        settings (Document, required): Microsoft 365 Settings.

	Returns:
	        None
	"""
	email = account.user_principal_name
	if settings.email_field and employee.meta.has_field(settings.email_field):
		frappe.db.set_value("Employee", employee.name, settings.email_field, email)
	if not settings.create_erpnext_user or employee.user_id:
		return
	if not frappe.db.exists("User", email):
		user = frappe.get_doc(
			{
				"doctype": "User",
				"email": email,
				"first_name": employee.first_name,
				"last_name": employee.last_name,
				"send_welcome_email": 0,
			}
		)
		user.insert(ignore_permissions=True)
	frappe.db.set_value("Employee", employee.name, "user_id", email)


def welcome(client: GraphClient, employee, account, settings, password: str) -> None:
	"""
	Send the welcome email when switched on.

	Parameters:
	        client (GraphClient, required): Graph client.
	        employee (Document, required): Employee.
	        account (Document, required): Microsoft 365 Account.
	        settings (Document, required): Microsoft 365 Settings.
	        password (str, required): Temporary password.

	Returns:
	        None
	"""
	if settings.send_welcome_email:
		recipient = send_welcome(client, employee, account.user_principal_name, password, settings)
		log("Welcome Email", account=account.name, details=_("Sent to {0}").format(recipient))

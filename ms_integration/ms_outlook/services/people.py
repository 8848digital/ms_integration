# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""
People for document groups: field values (Users, Employees, emails) turned
into email addresses and then into Microsoft 365 user ids.
"""

import frappe

from ms_integration.ms_core.services.field_paths import values_at
from ms_integration.ms_outlook.services.directory_api import find_user_by_address

ACCOUNT = "Microsoft 365 Account"


def addresses(doc, paths: list[str]) -> list[str]:
	"""
	Distinct email addresses behind field paths.

	Parameters:
	        doc (Document, required): Document.
	        paths (list[str], required): Field paths.

	Returns:
	        list[str]: Lower-case addresses, in order.
	"""
	found = []
	for path in paths:
		for value in values_at(doc, path):
			email = address_for(str(value))
			if email and email not in found:
				found.append(email)
	return found


def address_for(value: str) -> str | None:
	"""
	Work address for an Employee, a User or a plain email.

	An employee with a Microsoft 365 account uses its sign-in name; otherwise the
	Employee's company email / user, or the User's email.

	Parameters:
	        value (str, required): Field value.

	Returns:
	        str | None: Address.
	"""
	employee = value if frappe.db.exists("Employee", value) else None
	if not employee and frappe.db.exists("User", value):
		employee = frappe.db.get_value("Employee", {"user_id": value})
		if not employee:
			return frappe.db.get_value("User", value, "email").lower()
	if employee:
		upn = frappe.db.get_value(
			ACCOUNT, {"employee": employee, "status": "Active"}, "user_principal_name"
		)
		row = frappe.db.get_value("Employee", employee, ["company_email", "user_id"], as_dict=1)
		email = upn or row.company_email or row.user_id
		return email.lower() if email else None
	return value.strip().lower() if "@" in value else None


class Resolver:
	"""Email → Microsoft 365 user id, remembering answers for one sync."""

	def __init__(self, client):
		"""
		Parameters:
		        client (GraphClient, required): Graph client.
		"""
		self.client, self.known = client, {}

	def user_id(self, email: str) -> str | None:
		"""
		Object id of the Microsoft 365 user with this address.

		Parameters:
		        email (str, required): Address.

		Returns:
		        str | None: Object id, or None when the address isn't in the tenant.
		"""
		if email not in self.known:
			account = frappe.db.get_value(
				ACCOUNT, {"user_principal_name": email, "status": "Active"}, "object_id"
			)
			user = None if account else find_user_by_address(self.client, email)
			self.known[email] = account or (user or {}).get("id")
		return self.known[email]

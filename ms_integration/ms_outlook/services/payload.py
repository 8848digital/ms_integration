# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Microsoft 365 user properties built from the Employee and the Field Mapping."""

import frappe
from frappe.utils import getdate

from ms_integration.ms_core.services.settings import field_value

EXTENSION = "extensionAttribute"


def profile(employee, settings) -> dict:
	"""
	User properties for the employee (names, usage location and mapped fields).

	Empty values are sent as null so a cleared ERPNext field clears Microsoft 365.

	Parameters:
	        employee (Document, required): Employee.
	        settings (Document, required): Microsoft 365 Settings.

	Returns:
	        dict: Graph user properties (no password, no sign-in name).
	"""
	first = (employee.first_name or "").strip()
	last = (employee.last_name or "").strip()
	values = {
		"givenName": first or None,
		"surname": last or None,
		"displayName": (employee.employee_name or f"{first} {last}").strip(),
		"usageLocation": (settings.usage_location or "").strip().upper() or None,
	}
	for row in settings.field_map or []:
		values[row.microsoft_field] = as_text(field_value(employee, row.employee_field))
	return shape(values)


def as_text(value) -> str | None:
	"""
	A field value as Graph text (None for empty).

	Parameters:
	        value (Any, optional): Field value.

	Returns:
	        str | None: Text.
	"""
	if value in (None, ""):
		return None
	return str(value).strip()


def shape(values: dict) -> dict:
	"""
	Move flat property names into Graph's shapes (extension attributes, phones, dates).

	Parameters:
	        values (dict, required): Flat {property: text}.

	Returns:
	        dict: Graph user body.
	"""
	body, extensions = {}, {}
	for key, value in values.items():
		if key.startswith(EXTENSION):
			extensions[key] = value
		elif key == "businessPhone":
			body["businessPhones"] = [value] if value else []
		elif key == "employeeHireDate":
			body[key] = f"{getdate(value).isoformat()}T00:00:00Z" if value else None
		else:
			body[key] = value
	if extensions:
		body["onPremisesExtensionAttributes"] = extensions
	return body


def changed(before: dict, after: dict) -> dict:
	"""
	Properties whose value differs (extension attributes compared one by one).

	Parameters:
	        before (dict, required): Previous profile.
	        after (dict, required): Current profile.

	Returns:
	        dict: Properties to PATCH.
	"""
	diff = {
		k: v for k, v in after.items() if k != "onPremisesExtensionAttributes" and before.get(k) != v
	}
	old_ext = before.get("onPremisesExtensionAttributes") or {}
	new_ext = {
		k: v
		for k, v in (after.get("onPremisesExtensionAttributes") or {}).items()
		if old_ext.get(k) != v
	}
	if new_ext:
		diff["onPremisesExtensionAttributes"] = new_ext
	return diff


def manager_object_id(employee, settings) -> str | None:
	"""
	Microsoft 365 user id of the employee's manager, when the manager has an account.

	Parameters:
	        employee (Document, required): Employee.
	        settings (Document, required): Microsoft 365 Settings.

	Returns:
	        str | None: Manager object id.
	"""
	manager = employee.get(settings.manager_field) if settings.manager_field else None
	if not manager:
		return None
	return frappe.db.get_value(
		"Microsoft 365 Account", {"employee": manager, "status": "Active"}, "object_id"
	)

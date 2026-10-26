# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Post-install defaults."""

import frappe

DEFAULT_FIELD_MAP = [
	("designation", "jobTitle"),
	("department", "department"),
	("cell_number", "mobilePhone"),
	("name", "employeeId"),
	("company", "companyName"),
]


def after_install() -> None:
	"""
	Fill a sensible default Field Mapping (standard Employee fields only).

	Returns:
	        None
	"""
	settings = frappe.get_single("Outlook Settings")
	if settings.field_map:
		return
	for employee_field, microsoft_field in DEFAULT_FIELD_MAP:
		settings.append(
			"field_map", {"employee_field": employee_field, "microsoft_field": microsoft_field}
		)
	settings.flags.ignore_mandatory = True
	settings.save(ignore_permissions=True)

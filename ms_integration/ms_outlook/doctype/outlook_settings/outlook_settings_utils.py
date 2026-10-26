# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Validation for Outlook Settings."""

import re

import frappe
from frappe import _

from ms_integration.ms_core.services.settings import validate_condition

TOKENS = ("{first}", "{middle}", "{last}", "{first_initial}")


def validate_outlook_settings(doc) -> None:
	"""
	Normalise and check the settings.

	Parameters:
	        doc (Document, required): Outlook Settings.

	Returns:
	        None
	"""
	validate_condition(doc.create_condition, _("Create When"))
	validate_condition(doc.disable_condition, _("Disable When"))
	doc.email_domain = (doc.email_domain or "").strip().lower().lstrip("@") or None
	if doc.email_domain and not re.fullmatch(r"[a-z0-9-]+(\.[a-z0-9-]+)+", doc.email_domain):
		frappe.throw(_("Email Domain must look like example.com"))
	doc.email_pattern = (doc.email_pattern or "{first}.{last}").strip().lower()
	if not any(token in doc.email_pattern for token in TOKENS):
		frappe.throw(_("Address Pattern needs at least one of {0}").format(", ".join(TOKENS)))
	doc.usage_location = (doc.usage_location or "").strip().upper() or None
	if doc.usage_location and not re.fullmatch(r"[A-Z]{2}", doc.usage_location):
		frappe.throw(_("Usage Location must be a two-letter country code, e.g. SA or IN"))
	for row in doc.field_map or []:
		row.employee_field = (row.employee_field or "").strip()

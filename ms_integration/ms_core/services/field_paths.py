# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""
Field paths used by rules: "field", "table.field" (child rows) and
"link.field" (the linked document's field), chained as deep as needed,
e.g. "employee.reports_to" or "users.user".
"""

import frappe
from frappe import _

TABLE_TYPES = ("Table", "Table MultiSelect")
STANDARD_FIELDS = ("name", "owner", "modified_by")


def values_at(doc, path: str) -> list:
	"""
	Values reached by a field path, following child tables and links.

	Parameters:
	        doc (Document | dict, required): Starting document.
	        path (str, required): Field path.

	Returns:
	        list: Non-empty values, in order (one per child row for tables).
	"""
	docs = [doc]
	parts = [part for part in (path or "").strip().split(".") if part]
	for index, fieldname in enumerate(parts):
		last = index == len(parts) - 1
		next_docs = []
		for current in docs:
			value = current.get(fieldname)
			field = (
				frappe.get_meta(current.doctype).get_field(fieldname) if current.get("doctype") else None
			)
			if last:
				next_docs.append(value)
			elif field and field.fieldtype in TABLE_TYPES:
				next_docs.extend(value or [])
			elif field and field.fieldtype == "Link" and value:
				next_docs.append(frappe.get_cached_doc(field.options, value))
		docs = next_docs
	return [value for value in docs if value not in (None, "")]


def first_value(doc, path: str):
	"""
	First value reached by a field path.

	Parameters:
	        doc (Document, required): Starting document.
	        path (str, required): Field path.

	Returns:
	        Any: The value, or None.
	"""
	values = values_at(doc, path)
	return values[0] if values else None


def validate_path(doctype: str, path: str) -> None:
	"""
	Every step must exist; only the last step may be a plain field. "name",
	"owner" and "modified_by" are allowed.

	Parameters:
	        doctype (str, required): Starting DocType.
	        path (str, required): Field path.

	Returns:
	        None
	"""
	parts = [part for part in (path or "").strip().split(".") if part]
	current = doctype
	for index, fieldname in enumerate(parts):
		if fieldname in STANDARD_FIELDS and index == len(parts) - 1:
			return
		field = frappe.get_meta(current).get_field(fieldname)
		if not field:
			frappe.throw(_("Field {0} does not exist on {1}").format(frappe.bold(fieldname), current))
		if index < len(parts) - 1:
			if field.fieldtype not in (*TABLE_TYPES, "Link"):
				frappe.throw(_("{0} on {1} is not a table or link").format(frappe.bold(fieldname), current))
			current = field.options

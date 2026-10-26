# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Microsoft 365 connection settings and small shared helpers (rule conditions, field paths)."""

import frappe

SETTINGS = "Microsoft 365 Settings"


def get_connection():
	"""
	The cached Microsoft 365 Settings (tenant, app credentials, Test Mode).

	Returns:
	        Document: Microsoft 365 Settings.
	"""
	return frappe.get_cached_doc(SETTINGS)


def is_connected() -> bool:
	"""
	Whether the Microsoft 365 connection is switched on.

	Returns:
	        bool: True when Enabled is ticked in Microsoft 365 Settings.
	"""
	return bool(get_connection().enabled)


def lines(text: str | None) -> list[str]:
	"""
	Non-empty, stripped lines of a Small Text setting.

	Parameters:
	        text (str, optional): Setting value.

	Returns:
	        list[str]: One entry per line.
	"""
	return [line.strip() for line in (text or "").splitlines() if line.strip()]


def matches(condition: str | None, doc) -> bool:
	"""
	Evaluate a rule condition on a document; empty means always.

	Conditions get `doc`, `nowdate` and the safe `frappe` namespace, as in
	Frappe's Notification.

	Parameters:
	        condition (str, optional): Python expression.
	        doc (Document, required): Document being tested.

	Returns:
	        bool: Result of the expression.
	"""
	if not (condition or "").strip():
		return True
	from frappe.utils import nowdate
	from frappe.utils.safe_exec import get_safe_globals

	context = {"doc": doc, "nowdate": nowdate, "frappe": get_safe_globals().get("frappe")}
	return bool(frappe.safe_eval(condition, None, context))


def validate_condition(condition: str | None, label: str) -> None:
	"""
	Reject a condition that isn't valid Python.

	Parameters:
	        condition (str, optional): Python expression.
	        label (str, required): Field label for the message.

	Returns:
	        None
	"""
	if not (condition or "").strip():
		return
	try:
		compile(condition, label, "eval")
	except SyntaxError as e:
		frappe.throw(frappe._("{0} is not a valid Python expression: {1}").format(label, e.msg))


def field_value(doc, path: str | None):
	"""
	Value of a field or a path through Link fields, e.g. `department.department_name`.

	Parameters:
	        doc (Document, required): Starting document.
	        path (str, optional): Field name or dotted path.

	Returns:
	        Any: The value, or None when a step is empty.
	"""
	if not path:
		return None
	first, _, rest = path.strip().partition(".")
	value = doc.get(first)
	if not rest or not value:
		return value
	link = doc.meta.get_field(first)
	if not link or link.fieldtype != "Link":
		return None
	target = frappe.get_doc(link.options, value)
	return field_value(target, rest)

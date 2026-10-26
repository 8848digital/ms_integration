# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Microsoft 365 Group Rules: which DocTypes have groups and which feed members."""

import frappe
from frappe import _

from ms_integration.ms_core.services.field_paths import validate_path
from ms_integration.ms_core.services.settings import lines, matches, validate_condition

RULE = "Microsoft 365 Group Rule"
CACHE_KEY = "microsoft365_group_rule_doctypes"


def rule_doctypes() -> dict:
	"""
	DocTypes with an enabled rule, and source DocTypes feeding members.

	Returns:
	        dict: {"targets": [doctype], "sources": {source: [[target, link_field]]}}.
	"""
	return frappe.cache.get_value(CACHE_KEY, generator=build_rule_doctypes)


def build_rule_doctypes() -> dict:
	"""
	Build the cached map from the enabled rules.

	Returns:
	        dict: See rule_doctypes.
	"""
	targets, sources = [], {}
	for name in frappe.get_all(RULE, {"enabled": 1}, pluck="name"):
		rule = frappe.get_doc(RULE, name)
		targets.append(rule.reference_doctype)
		for row in rule.member_sources:
			sources.setdefault(row.source_doctype, []).append([rule.reference_doctype, row.link_field])
	return {"targets": targets, "sources": sources}


def clear_cache() -> None:
	"""
	Forget the cached map (after a rule changes).

	Returns:
	        None
	"""
	frappe.cache.delete_value(CACHE_KEY)


def rule_for(doctype: str):
	"""
	The enabled rule of a DocType.

	Parameters:
	        doctype (str, required): DocType.

	Returns:
	        Document | None: Microsoft 365 Group Rule.
	"""
	if frappe.db.get_value(RULE, doctype, "enabled"):
		return frappe.get_cached_doc(RULE, doctype)
	return None


def parents_of(source) -> list[tuple[str, str]]:
	"""
	Documents whose group gets members from this source document.

	Parameters:
	        source (Document, required): Source document (e.g. an Email Group Member).

	Returns:
	        list[tuple[str, str]]: (DocType, name) pairs.
	"""
	pairs = []
	for target, link_field in rule_doctypes()["sources"].get(source.doctype, []):
		value = source.get(link_field)
		if value and (target, value) not in pairs:
			pairs.append((target, value))
	return pairs


def source_members(rule, doc) -> list[str]:
	"""
	Addresses from the rule's linked-document sources.

	Parameters:
	        rule (Document, required): Microsoft 365 Group Rule.
	        doc (Document, required): The rule's document.

	Returns:
	        list[str]: Addresses.
	"""
	from ms_integration.ms_outlook.services.people import addresses

	found = []
	for row in rule.member_sources:
		names = frappe.get_all(row.source_doctype, {row.link_field: doc.name}, pluck="name", limit=1000)
		for name in names:
			source = frappe.get_doc(row.source_doctype, name)
			if matches(row.source_condition, source):
				found += [a for a in addresses(source, [row.member_field]) if a not in found]
	return found


def validate_rule(rule) -> None:
	"""
	Check conditions, field paths, sources and write-back fields.

	Parameters:
	        rule (Document, required): Microsoft 365 Group Rule.

	Returns:
	        None
	"""
	validate_condition(rule.condition, _("Create When"))
	validate_condition(rule.close_condition, _("Close When"))
	for path in lines(rule.member_fields) + lines(rule.owner_fields):
		validate_path(rule.reference_doctype, path)
	for row in rule.member_sources:
		validate_condition(row.source_condition, _("Only When"))
		link = frappe.get_meta(row.source_doctype).get_field(row.link_field)
		if not link or link.fieldtype != "Link" or link.options != rule.reference_doctype:
			frappe.throw(
				_("{0} on {1} must be a Link to {2}").format(
					row.link_field, row.source_doctype, rule.reference_doctype
				)
			)
		validate_path(row.source_doctype, row.member_field)
	meta = frappe.get_meta(rule.reference_doctype)
	for field in (rule.group_id_field, rule.group_email_field):
		if field and not meta.has_field(field):
			frappe.throw(_("Field {0} does not exist on {1}").format(field, rule.reference_doctype))

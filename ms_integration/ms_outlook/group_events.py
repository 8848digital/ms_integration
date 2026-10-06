# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""doc_events "*": keep document groups in line with any DocType that has a Group Rule."""

import frappe

from ms_integration.ms_outlook.services.document_group_close import document_deleted
from ms_integration.ms_outlook.services.group_rules import parents_of, rule_doctypes
from ms_integration.ms_outlook.services.settings import is_enabled
from ms_integration.ms_outlook.tasks import enqueue

SKIP = ("Microsoft 365 Document Group", "Microsoft 365 Log", "Microsoft 365 Account")


def on_update(doc, method=None):
	"""
	Queue a group sync for the document (when it has a rule) and for the
	documents it feeds members to. Exits after a cached lookup otherwise.

	Parameters:
	        doc (Document, required): Saved document.
	        method (str, optional): Hook name.

	Returns:
	        None
	"""
	for doctype, name in affected(doc):
		enqueue("sync_document_group", f"{doctype}:{name}", doctype=doctype, name=name)


def on_trash(doc, method=None):
	"""
	A document about to be deleted drops its group record (and its group, if
	the rule says so).

	Parameters:
	        doc (Document, required): Document being deleted.
	        method (str, optional): Hook name.

	Returns:
	        None
	"""
	if doc.doctype in SKIP or frappe.flags.in_install or not is_enabled():
		return
	if doc.doctype in rule_doctypes()["targets"]:
		group_id = document_deleted(doc)
		if group_id:
			enqueue("delete_group", group_id, group_id=group_id)


def after_delete(doc, method=None):
	"""
	A deleted source document (e.g. an Email Group Member) resyncs the groups it fed.

	Parameters:
	        doc (Document, required): Deleted document.
	        method (str, optional): Hook name.

	Returns:
	        None
	"""
	if doc.doctype in SKIP or frappe.flags.in_install or not is_enabled():
		return
	for doctype, name in parents_of(doc):
		if frappe.db.exists(doctype, name):
			enqueue("sync_document_group", f"{doctype}:{name}", doctype=doctype, name=name)


def affected(doc) -> list[tuple[str, str]]:
	"""
	(DocType, name) pairs whose group should be synced after this save.

	Parameters:
	        doc (Document, required): Saved document.

	Returns:
	        list[tuple[str, str]]: Pairs.
	"""
	if doc.doctype in SKIP or frappe.flags.in_install or frappe.flags.in_migrate:
		return []
	kinds = rule_doctypes()
	if doc.doctype not in kinds["targets"] and doc.doctype not in kinds["sources"]:
		return []
	if not is_enabled():
		return []
	pairs = [(doc.doctype, doc.name)] if doc.doctype in kinds["targets"] else []
	return pairs + [pair for pair in parents_of(doc) if pair not in pairs]

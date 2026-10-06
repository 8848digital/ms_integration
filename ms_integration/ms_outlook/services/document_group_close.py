# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Closing a document group and dropping it when its document is deleted."""

import frappe
from frappe import _
from frappe.utils import now_datetime

from ms_integration.ms_core.services.graph import GraphClient, friendly
from ms_integration.ms_core.services.logs import log
from ms_integration.ms_outlook.services import directory_api as api
from ms_integration.ms_outlook.services import groups_api
from ms_integration.ms_outlook.services.group_rules import rule_for

RECORD = "Microsoft 365 Document Group"


def close(client: GraphClient, record, rule) -> None:
	"""
	The document closed: remove members, delete the group, or only mark it closed.

	Parameters:
	        client (GraphClient, required): Graph client.
	        record (Document, required): Microsoft 365 Document Group.
	        rule (Document, required): Microsoft 365 Group Rule.

	Returns:
	        None
	"""
	try:
		if rule.on_close == "Delete Group":
			groups_api.delete_group(client, record.group_id)
		elif rule.on_close == "Remove Members":
			owners = {m["id"] for m in groups_api.group_owners(client, record.group_id)}
			for member in groups_api.group_members(client, record.group_id):
				if member["id"] not in owners:
					api.remove_from_group(client, record.group_id, member["id"])
	except Exception as e:
		record.db_set("error", friendly(e))
		log("Document Group", False, details=f"Close {record.display_name}", error=e)
		return
	status = "Deleted" if rule.on_close == "Delete Group" else "Closed"
	record.db_set({"status": status, "error": None, "last_synced": now_datetime()})
	log(
		"Document Group", details=_("{0}: {1} ({2})").format(record.display_name, status, rule.on_close)
	)


def document_deleted(doc) -> str | None:
	"""
	Before a document is deleted: drop its group record; return the group id
	to delete when the rule says so.

	Parameters:
	        doc (Document, required): Document being deleted.

	Returns:
	        str | None: Group id to delete, or None.
	"""
	found = frappe.db.get_value(
		RECORD, {"reference_doctype": doc.doctype, "reference_name": doc.name}
	)
	record = frappe.get_doc(RECORD, found) if found else None
	if not record:
		return None
	rule = rule_for(doc.doctype)
	group_id = (
		record.group_id
		if rule and rule.on_delete == "Delete Group" and record.status != "Deleted"
		else None
	)
	frappe.delete_doc(RECORD, record.name, ignore_permissions=True, force=True)
	return group_id

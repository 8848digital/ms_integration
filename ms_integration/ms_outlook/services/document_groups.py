# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""
A Microsoft 365 group per document: created when the rule's condition matches,
members kept in line with the document's fields and linked documents, members
removed or the group deleted when the document closes.
"""

import frappe
from frappe import _
from frappe.utils import now_datetime

from ms_integration.ms_core.services.graph import GraphClient, friendly
from ms_integration.ms_core.services.logs import log, summary
from ms_integration.ms_core.services.settings import lines, matches
from ms_integration.ms_outlook.services import directory_api as api
from ms_integration.ms_outlook.services import groups_api
from ms_integration.ms_outlook.services.directory import upsert_group
from ms_integration.ms_outlook.services.document_group_close import close
from ms_integration.ms_outlook.services.group_rules import rule_for, source_members
from ms_integration.ms_outlook.services.people import Resolver, addresses

RECORD = "Microsoft 365 Document Group"


def record_for(doctype: str, name: str):
	"""
	The document's group record, if any.

	Parameters:
	        doctype (str, required): DocType.
	        name (str, required): Document name.

	Returns:
	        Document | None: Microsoft 365 Document Group.
	"""
	found = frappe.db.get_value(RECORD, {"reference_doctype": doctype, "reference_name": name})
	return frappe.get_doc(RECORD, found) if found else None


def sync_document(doctype: str, name: str) -> None:
	"""
	Create, update or close the document's group as its rule says.

	Parameters:
	        doctype (str, required): DocType.
	        name (str, required): Document name.

	Returns:
	        None
	"""
	rule = rule_for(doctype)
	if not rule or not frappe.db.exists(doctype, name):
		return
	doc = frappe.get_doc(doctype, name)
	record = record_for(doctype, name)
	client = GraphClient()
	if (rule.close_condition or "").strip() and matches(rule.close_condition, doc):
		if record and record.status == "Active":
			close(client, record, rule)
		return
	if not matches(rule.condition, doc):
		return
	record = record or new_record(doctype, name)
	try:
		ensure_group(client, record, rule, doc)
		sync_members(client, record, rule, doc)
	except Exception as e:
		record.db_set(
			{"status": "Failed" if not record.group_id else record.status, "error": friendly(e)}
		)
		log("Document Group", False, details=f"{doctype} {name}", error=e)


def new_record(doctype: str, name: str):
	"""
	A Queued group record for the document.

	Parameters:
	        doctype (str, required): DocType.
	        name (str, required): Document name.

	Returns:
	        Document: Microsoft 365 Document Group.
	"""
	record = frappe.get_doc({"doctype": RECORD, "reference_doctype": doctype, "reference_name": name})
	return record.insert(ignore_permissions=True)


def ensure_group(client: GraphClient, record, rule, doc) -> None:
	"""
	Create the group when the record has none (or it was deleted).

	Parameters:
	        client (GraphClient, required): Graph client.
	        record (Document, required): Microsoft 365 Document Group.
	        rule (Document, required): Microsoft 365 Group Rule.
	        doc (Document, required): The rule's document.

	Returns:
	        None
	"""
	if record.group_id and record.status != "Deleted":
		return
	context = {"doc": doc}
	# Templates come from Group Rules (System Manager / HR Manager only).
	name = frappe.render_template(rule.name_template, context).strip()[:256]  # nosemgrep
	template = rule.description_template or ""
	description = frappe.render_template(template, context).strip()  # nosemgrep
	resolver = Resolver(client)
	owners = [i for i in (resolver.user_id(a) for a in addresses(doc, lines(rule.owner_fields))) if i]
	group = groups_api.create_group(
		client, name, rule.group_type, description, rule.visibility, owners[:20]
	)
	upsert_group(group)
	record.update(
		{
			"group": group["id"],
			"group_id": group["id"],
			"display_name": name,
			"mail": group.get("mail"),
			"status": "Active",
			"error": None,
		}
	)
	record.save(ignore_permissions=True)
	write_back(rule, doc, group)
	log(
		"Document Group", details=_("Created group {0} for {1} {2}").format(name, doc.doctype, doc.name)
	)


def write_back(rule, doc, group: dict) -> None:
	"""
	Save the group id / email on the document when configured.

	Parameters:
	        rule (Document, required): Microsoft 365 Group Rule.
	        doc (Document, required): The rule's document.
	        group (dict, required): Graph group.

	Returns:
	        None
	"""
	values = {}
	if rule.group_id_field:
		values[rule.group_id_field] = group["id"]
	if rule.group_email_field:
		values[rule.group_email_field] = group.get("mail")
	if values:
		frappe.db.set_value(doc.doctype, doc.name, values, update_modified=False)


def sync_members(client: GraphClient, record, rule, doc) -> None:
	"""
	Add missing members and owners; remove members no longer listed (if set).

	Parameters:
	        client (GraphClient, required): Graph client.
	        record (Document, required): Microsoft 365 Document Group (Active).
	        rule (Document, required): Microsoft 365 Group Rule.
	        doc (Document, required): The rule's document.

	Returns:
	        None
	"""
	resolver = Resolver(client)
	owners = addresses(doc, lines(rule.owner_fields))
	wanted = owners + [
		a
		for a in addresses(doc, lines(rule.member_fields)) + source_members(rule, doc)
		if a not in owners
	]
	current = {m["id"] for m in groups_api.group_members(client, record.group_id)}
	current_owners = {m["id"] for m in groups_api.group_owners(client, record.group_id)}
	rows, added, removed, wanted_ids = [], [], [], set()
	for email in wanted:
		row = {"email": email, "role": "Owner" if email in owners else "Member", "status": "Added"}
		user_id = resolver.user_id(email)
		if not user_id:
			row.update(status="Not in Microsoft 365", note=_("No Microsoft 365 user with this address"))
			rows.append(row)
			continue
		wanted_ids.add(user_id)
		try:
			if user_id not in current:
				api.add_to_group(client, record.group_id, user_id)
				added.append(email)
			if row["role"] == "Owner" and user_id not in current_owners:
				groups_api.add_owner(client, record.group_id, user_id)
		except Exception as e:
			row.update(status="Failed", note=friendly(e)[:140])
		rows.append(row)
	if rule.remove_missing_members:
		for user_id in current - wanted_ids - current_owners:
			api.remove_from_group(client, record.group_id, user_id)
			removed.append(user_id)
	record.set("members", rows)
	record.update({"status": "Active", "last_synced": now_datetime(), "error": None})
	record.save(ignore_permissions=True)
	if added or removed:
		log("Document Group", details=f"{record.display_name}: " + summary(added, removed, "members"))

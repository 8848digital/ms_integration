# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Document group actions for the form button of DocTypes with a Group Rule."""

import frappe
from frappe import _

from ms_integration.ms_outlook.services.document_groups import record_for
from ms_integration.ms_outlook.services.group_rules import rule_for
from ms_integration.ms_outlook.services.settings import is_enabled
from ms_integration.ms_outlook.tasks import enqueue
from ms_integration.utils.api_handlers.response_formatter import api_response


@frappe.whitelist(methods=["GET"])
def get_group(doctype: str, name: str):
	"""
	        The document's Microsoft 365 group, for the form button.

	        **Endpoint:** `/api/method/ms_integration.ms_outlook.api.v1.group.get_group`
	        **HTTP Method:** GET
	        **Parameters:**
	                - doctype (str, required): DocType of the document
	                - name (str, required): Document name
	        **Response:**
	```json
	                {"status": true, "status_code": 200, "message": "Group found",
	                 "data": {"has_rule": true, "record": "Project-PROJ-0001", "status": "Active",
	                          "display_name": "PRJ - Website", "mail": "prj-website@example.com"}, "errors": null}
	```
	"""
	frappe.has_permission(doctype, "read", name, throw=True)
	data = {"has_rule": bool(is_enabled() and rule_for(doctype))}
	record = record_for(doctype, name) if data["has_rule"] else None
	if record and frappe.has_permission("Microsoft 365 Document Group", "read"):
		data.update(
			record=record.name, status=record.status, display_name=record.display_name, mail=record.mail
		)
	return api_response(message=_("Group found") if record else _("No group"), data=data)


@frappe.whitelist(methods=["POST"])
def sync_group(doctype: str, name: str):
	"""
	        Create the document's group now, or re-sync its members.

	        **Endpoint:** `/api/method/ms_integration.ms_outlook.api.v1.group.sync_group`
	        **HTTP Method:** POST
	        **Parameters:**
	                - doctype (str, required): DocType of the document
	                - name (str, required): Document name
	        **Response:**
	```json
	                {"status": true, "status_code": 200, "message": "Syncing the Microsoft 365 group…",
	                 "data": null, "errors": null}
	```
	"""
	frappe.has_permission(doctype, "write", name, throw=True)
	frappe.has_permission("Microsoft 365 Document Group", "write", throw=True)
	if not (is_enabled() and rule_for(doctype)):
		frappe.throw(_("No enabled Microsoft 365 Group Rule for {0}.").format(_(doctype)))
	enqueue("sync_document_group", f"{doctype}:{name}", doctype=doctype, name=name)
	return api_response(message=_("Syncing the Microsoft 365 group…"))

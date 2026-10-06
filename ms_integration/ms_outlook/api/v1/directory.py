# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Outlook directory sync action."""

import frappe
from frappe import _

from ms_integration.ms_outlook.services.settings import is_enabled
from ms_integration.ms_outlook.tasks import enqueue
from ms_integration.utils.api_handlers.response_formatter import api_response


@frappe.whitelist(methods=["POST"])
def sync_directory():
	"""
	        Queue a directory sync: licenses, groups, every account's licenses and
	        groups, and linking employees who already have an account.

	        **Endpoint:** `/api/method/ms_integration.ms_core.api.v1.settings.sync_directory`
	        **HTTP Method:** POST
	        **Parameters:** none
	        **Response:**
	```json
	                {"status": true, "status_code": 200, "message": "Directory sync started…",
	                 "data": null, "errors": null}
	```
	"""
	frappe.only_for(("System Manager", "HR Manager"))
	if not is_enabled():
		frappe.throw(
			_("Outlook integration is not enabled (Outlook Settings and Microsoft 365 Settings).")
		)
	enqueue("sync_directory", "all")
	return api_response(message=_("Directory sync started. Licenses and groups appear in a minute."))

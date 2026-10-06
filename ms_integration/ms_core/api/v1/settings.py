# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Microsoft 365 Settings actions: test the connection."""

import frappe
from frappe import _

from ms_integration.ms_core.services.graph import GraphClient, friendly
from ms_integration.ms_core.services.logs import log
from ms_integration.utils.api_handlers.response_formatter import api_response


@frappe.whitelist(methods=["POST"])
def test_connection():
	"""
	        Get a token and read the tenant; reports the tenant name and whether the
	        Email Domain is one of its verified domains.

	        **Endpoint:** `/api/method/ms_integration.ms_core.api.v1.settings.test_connection`
	        **HTTP Method:** POST
	        **Parameters:** none
	        **Response:**
	```json
	                {"status": true, "status_code": 200, "message": "Connected to Contoso",
	                 "data": {"tenant": "Contoso", "domains": ["contoso.com"], "domain_verified": true},
	                 "errors": null}
	```
	"""
	frappe.only_for("System Manager")
	try:
		org = (
			GraphClient()
			.get("/organization", params={"$select": "displayName,verifiedDomains"})
			.get("value")
			or [{}]
		)[0]
	except Exception as e:
		log("Connection", False, error=e)
		return api_response(status=False, code=200, message=friendly(e))
	domains = [d["name"].lower() for d in org.get("verifiedDomains") or []]
	# Outlook's email domain, when the Outlook module is set up (checked by name, no import).
	domain = (frappe.db.get_single_value("Outlook Settings", "email_domain") or "").strip().lower()
	data = {
		"tenant": org.get("displayName"),
		"domains": domains,
		"domain_verified": domain in domains,
	}
	message = _("Connected to {0}").format(data["tenant"])
	if domain and not data["domain_verified"]:
		message += ". " + _("Warning: {0} is not a verified domain of this tenant.").format(domain)
	log("Connection", details=message)
	return api_response(message=message, data=data)

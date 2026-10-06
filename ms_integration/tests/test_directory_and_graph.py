# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Directory sync, the real Graph client (HTTP mocked) and API permissions."""

from unittest.mock import MagicMock, patch

import frappe
from frappe.tests import IntegrationTestCase

from ms_integration.ms_core.services import test_mode
from ms_integration.ms_core.services.graph import GraphClient, GraphError, quote
from ms_integration.ms_outlook.api.v1 import account as account_api
from ms_integration.ms_outlook.services.directory import sync_all
from ms_integration.tests.m365_fixtures import DOMAIN, make_employee, setup_microsoft365

HTTP = "ms_integration.ms_core.services.graph.requests"


def reply(status: int = 200, body: dict | None = None, headers: dict | None = None):
	"""
	A fake requests.Response.

	Parameters:
	        status (int, optional): HTTP status.
	        body (dict, optional): JSON body.
	        headers (dict, optional): Headers.

	Returns:
	        MagicMock: Response.
	"""
	response = MagicMock(
		status_code=status, headers=headers or {}, content=b"{}" if body is not None else b""
	)
	response.json.return_value = body or {}
	response.text = frappe.as_json(body or {})
	return response


class TestDirectoryAndGraph(IntegrationTestCase):
	"""Directory sync in Test Mode; Graph client behaviour against mocked HTTP."""

	def setUp(self):
		"""
		Fresh Test Mode tenant.

		Returns:
		        None
		"""
		setup_microsoft365()

	def tearDown(self):
		"""
		Stop patches; back to Administrator.

		Returns:
		        None
		"""
		patch.stopall()
		frappe.set_user("Administrator")

	def test_sync_licenses_groups_and_link_existing(self):
		"""Seats counted, group types and manageability set, existing mailbox linked."""
		frappe.db.set_single_value("Outlook Settings", "auto_create", 0)
		employee = make_employee("It", "Admin", company_email=f"it.admin@{DOMAIN}")
		counts = sync_all()
		self.assertEqual((counts["licenses"], counts["groups"], counts["linked"]), (3, 5, 1))
		self.assertEqual(
			frappe.db.get_value("Microsoft 365 License", "EXCHANGESTANDARD", "available_units"), 2
		)
		groups = {
			g.display_name: g
			for g in frappe.get_all(
				"Microsoft 365 Group", ["display_name", "group_type", "membership_managed"]
			)
		}
		self.assertEqual(groups["Sales Team"].group_type, "Distribution")
		self.assertEqual(groups["IT Team"].group_type, "Security")
		self.assertFalse(groups["Everyone (Dynamic)"].membership_managed)
		self.assertTrue(groups["All Staff"].membership_managed)
		self.assertEqual(frappe.db.get_value("Microsoft 365 Account", employee.name, "status"), "Active")

	def test_token_cached_retry_on_throttle_and_paging(self):
		"""One token for many calls; 429 waits Retry-After; nextLink pages are followed."""
		frappe.db.set_single_value(
			"Microsoft 365 Settings", {"test_mode": 0, "tenant_id": "t", "client_id": "c"}
		)
		frappe.cache.delete_keys("microsoft365_token")
		http = patch(HTTP).start()
		sleep = patch("ms_integration.ms_core.services.graph.time.sleep").start()
		http.post.return_value = reply(body={"access_token": "tok", "expires_in": 3600})
		http.request.side_effect = [
			reply(
				429, {"error": {"code": "TooManyRequests", "message": "slow down"}}, {"Retry-After": "2"}
			),
			reply(
				body={"value": [1], "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=x"}
			),
			reply(body={"value": [2]}),
		]
		self.assertEqual(GraphClient().get_all("/users"), [1, 2])
		self.assertEqual(http.post.call_count, 1)
		sleep.assert_called_once_with(2)
		self.assertEqual(http.request.call_args.kwargs["headers"]["Authorization"], "Bearer tok")

	def test_errors_and_filter_quoting(self):
		"""Graph errors carry Microsoft's code and message; quotes are escaped in filters."""
		frappe.db.set_single_value("Microsoft 365 Settings", "test_mode", 0)
		frappe.cache.delete_keys("microsoft365_token")
		http = patch(HTTP).start()
		http.post.return_value = reply(
			400,
			{
				"error": "invalid_client",
				"error_description": "AADSTS7000215: Invalid client secret.\r\nTrace",
			},
		)
		with self.assertRaises(GraphError) as caught:
			GraphClient().get("/organization")
		self.assertEqual(caught.exception.code, "invalid_client")
		self.assertNotIn("Trace", caught.exception.message)
		self.assertEqual(quote("o'brien@x.com"), "o''brien@x.com")

	def test_api_needs_permission(self):
		"""A user without Microsoft 365 Account create permission can't create accounts."""
		employee = make_employee("Zoe", "Lim", status="Inactive")
		if not frappe.db.exists("User", "no-m365@example.com"):
			frappe.get_doc({"doctype": "User", "email": "no-m365@example.com", "first_name": "No"}).insert()
		frappe.set_user("no-m365@example.com")
		self.assertRaises(frappe.PermissionError, account_api.create_account, employee.name)
		frappe.set_user("Administrator")
		reply_ = account_api.create_account(employee.name)
		self.assertTrue(reply_["status"])
		self.assertEqual(frappe.db.get_value("Microsoft 365 Account", employee.name, "status"), "Active")
		self.assertEqual(account_api.get_status(employee.name)["data"]["email"], f"zoe.l@{DOMAIN}")
		self.assertEqual(len(test_mode.load()["users"]), 2)

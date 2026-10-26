# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Phase 2: a Microsoft 365 group per document (Project, Email Group) via Group Rules."""

from unittest.mock import patch

import frappe
from frappe.tests import IntegrationTestCase

from ms_integration.ms_core.services import test_mode
from ms_integration.tests.m365_fixtures import COMPANY, DOMAIN, make_employee, setup_microsoft365

RULE = "Microsoft 365 Group Rule"
RECORD = "Microsoft 365 Document Group"


def tenant_group(group_id: str) -> dict | None:
	"""
	A fake-tenant group by id.

	Parameters:
	        group_id (str, required): Group id.

	Returns:
	        dict | None: Stored group.
	"""
	return test_mode.load()["groups"].get(group_id)


def member_emails(group_id: str) -> set[str]:
	"""
	Sign-in names of a fake-tenant group's members.

	Parameters:
	        group_id (str, required): Group id.

	Returns:
	        set[str]: Addresses.
	"""
	tenant = test_mode.load()
	return {tenant["users"][i]["userPrincipalName"] for i in tenant["groups"][group_id]["members"]}


class TestDocumentGroups(IntegrationTestCase):
	"""Groups follow their documents: created, members synced, closed, deleted."""

	def setUp(self):
		"""
		Fresh tenant; two employees with accounts and ERPNext users.

		Returns:
		        None
		"""
		setup_microsoft365(create_erpnext_user=1)
		for doctype in (RULE, RECORD):
			frappe.db.delete(doctype)
		frappe.db.delete("Project", {"project_name": ("like", "M365 %")})
		frappe.db.delete("Email Group", {"title": ("like", "M365 %")})
		frappe.cache.delete_value("microsoft365_group_rule_doctypes")
		self.anna = f"anna.b@{DOMAIN}"
		self.ben = f"ben.c@{DOMAIN}"
		make_employee("Anna", "Bell")
		make_employee("Ben", "Cole")

	def tearDown(self):
		"""
		Stop patches.

		Returns:
		        None
		"""
		patch.stopall()

	def test_project_group_follows_users_and_closes(self):
		"""Project users become members; removed users leave; Completed removes members."""
		frappe.get_doc(
			{
				"doctype": RULE,
				"reference_doctype": "Project",
				"condition": 'doc.status == "Open"',
				"group_type": "Microsoft 365",
				"name_template": "PRJ - {{ doc.project_name }}",
				"member_fields": "users.user",
				"close_condition": 'doc.status == "Completed"',
				"on_close": "Remove Members",
			}
		).insert()
		project = frappe.get_doc(
			{
				"doctype": "Project",
				"company": COMPANY,
				"project_name": "M365 Website",
				"users": [
					{"user": self.anna, "welcome_email_sent": 1},
					{"user": self.ben, "welcome_email_sent": 1},
				],
			}
		).insert()
		record = frappe.get_doc(RECORD, f"Project-{project.name}")
		self.assertEqual((record.status, record.display_name), ("Active", "PRJ - M365 Website"))
		group = tenant_group(record.group_id)
		self.assertEqual((group["groupTypes"], group["mailNickname"]), (["Unified"], "prj-m365-website"))
		self.assertEqual(member_emails(record.group_id), {self.anna, self.ben})
		self.assertTrue(frappe.db.exists("Microsoft 365 Group", record.group_id))
		project.reload()
		project.users = [row for row in project.users if row.user == self.anna]
		project.save()
		self.assertEqual(member_emails(record.group_id), {self.anna})
		project.reload()
		project.status = "Completed"
		project.save()
		self.assertEqual(member_emails(record.group_id), set())
		self.assertEqual(frappe.db.get_value(RECORD, record.name, "status"), "Closed")

	def test_email_group_members_unsubscribe_and_delete(self):
		"""Email Group Members sync; outsiders flagged; unsubscribe/delete remove; group deleted with the list."""
		frappe.get_doc(
			{
				"doctype": RULE,
				"reference_doctype": "Email Group",
				"group_type": "Security",
				"name_template": "{{ doc.title }}",
				"on_delete": "Delete Group",
				"member_sources": [
					{
						"source_doctype": "Email Group Member",
						"link_field": "email_group",
						"member_field": "email",
						"source_condition": "not doc.unsubscribed",
					}
				],
			}
		).insert()
		email_group = frappe.get_doc({"doctype": "Email Group", "title": "M365 Newsletter"}).insert()
		members = {}
		for email in (self.anna, self.ben, "outsider@gmail.test"):
			members[email] = frappe.get_doc(
				{"doctype": "Email Group Member", "email_group": email_group.name, "email": email}
			).insert()
		record = frappe.get_doc(RECORD, f"Email Group-{email_group.name}")
		self.assertEqual(member_emails(record.group_id), {self.anna, self.ben})
		self.assertEqual(tenant_group(record.group_id)["securityEnabled"], True)
		statuses = {row.email: row.status for row in record.members}
		self.assertEqual(statuses["outsider@gmail.test"], "Not in Microsoft 365")
		members[self.ben].unsubscribed = 1
		members[self.ben].save()
		self.assertEqual(member_emails(record.group_id), {self.anna})
		members[self.anna].delete()
		self.assertEqual(member_emails(record.group_id), set())
		group_id = record.group_id
		frappe.db.delete("Email Group Member", {"email_group": email_group.name})
		email_group.delete()
		self.assertIsNone(tenant_group(group_id))
		self.assertFalse(frappe.db.exists(RECORD, record.name))

	def test_condition_and_validation(self):
		"""No group before the condition matches; a source link that doesn't point here is refused."""
		frappe.get_doc(
			{
				"doctype": RULE,
				"reference_doctype": "Project",
				"condition": 'doc.status == "Open" and doc.priority == "High"',
				"name_template": "{{ doc.name }}",
				"owner_fields": "users.user",
			}
		).insert()
		project = frappe.get_doc(
			{
				"doctype": "Project",
				"company": COMPANY,
				"project_name": "M365 Low",
				"priority": "Low",
				"users": [{"user": self.anna, "welcome_email_sent": 1}],
			}
		).insert()
		self.assertFalse(frappe.db.exists(RECORD, f"Project-{project.name}"))
		project.priority = "High"
		project.save()
		record = frappe.get_doc(RECORD, f"Project-{project.name}")
		anna_id = frappe.db.get_value(
			"Microsoft 365 Account", {"user_principal_name": self.anna}, "object_id"
		)
		self.assertIn(anna_id, tenant_group(record.group_id)["owners"])
		self.assertEqual(record.members[0].role, "Owner")
		bad = frappe.get_doc(
			{
				"doctype": RULE,
				"reference_doctype": "Email Group",
				"name_template": "x",
				"member_sources": [
					{"source_doctype": "ToDo", "link_field": "allocated_to", "member_field": "allocated_to"}
				],
			}
		)
		self.assertRaises(frappe.ValidationError, bad.insert)

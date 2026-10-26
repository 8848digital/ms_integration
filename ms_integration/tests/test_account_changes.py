# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Editing an account's licenses/groups, employee updates, offboarding and rehire."""

from unittest.mock import patch

import frappe
from frappe.tests import IntegrationTestCase
from frappe.utils import add_days, now_datetime

from ms_integration.ms_core.services import test_mode
from ms_integration.ms_outlook.services import lifecycle
from ms_integration.tests.m365_fixtures import (
	BUSINESS,
	DOMAIN,
	E3,
	EXCHANGE,
	group,
	make_employee,
	make_rule,
	setup_microsoft365,
	tenant_groups,
	tenant_licenses,
	tenant_user,
)


class TestAccountChanges(IntegrationTestCase):
	"""Changes made in ERPNext reach Microsoft 365 and failures never leave wrong rows."""

	def setUp(self):
		"""
		Fresh tenant; one employee with E3 + All Staff + IT Team.

		Returns:
		        None
		"""
		setup_microsoft365()
		make_rule("Everyone", licenses=[E3], groups=["All Staff", "IT Team"])
		self.employee = make_employee("Hana", "Saleh", designation=None)
		self.email = f"hana.s@{DOMAIN}"
		self.account = frappe.get_doc("Microsoft 365 Account", self.employee.name)

	def tearDown(self):
		"""
		Stop patches.

		Returns:
		        None
		"""
		patch.stopall()

	def test_edit_rows_adds_and_removes_in_microsoft(self):
		"""Swap a license and a group by editing the account rows."""
		self.account.set("licenses", [{"license": BUSINESS}])
		self.account.set("groups", [{"group": group("All Staff")}, {"group": group("Egypt Office")}])
		self.account.save()
		self.assertEqual(tenant_licenses(self.email), {BUSINESS})
		self.assertEqual(tenant_groups(self.email), {"All Staff", "Egypt Office"})
		self.assertTrue(frappe.db.exists("Microsoft 365 Log", {"action": "Groups", "status": "Success"}))

	def test_failed_license_change_restores_rows(self):
		"""No seats left: the license row goes back to what Microsoft 365 really has."""
		for i in range(2):
			tenant = test_mode.load()
			user = test_mode.new_user(
				tenant, {"userPrincipalName": f"x{i}@{DOMAIN}", "usageLocation": "SA"}
			)
			user["licenses"].append("4b9405b0-7788-4568-add1-99614e613b69")
			test_mode.save(tenant)
		self.account.append("licenses", {"license": EXCHANGE})
		self.account.save()
		self.assertEqual({r.license for r in self.account.licenses}, {E3})
		self.assertEqual(tenant_licenses(self.email), {E3})
		self.assertTrue(
			frappe.db.exists("Microsoft 365 Log", {"action": "Licenses", "status": "Failed"})
		)

	def test_distribution_list_cannot_be_added(self):
		"""Groups Graph can't change members of are refused with a clear message."""
		self.account.append("groups", {"group": group("Sales Team")})
		self.assertRaises(frappe.ValidationError, self.account.save)

	def test_employee_change_updates_microsoft(self):
		"""A new designation becomes the job title."""
		if not frappe.db.exists("Designation", "Analyst"):
			frappe.get_doc({"doctype": "Designation", "designation_name": "Analyst"}).insert()
		self.employee.reload()
		self.employee.designation = "Analyst"
		self.employee.save()
		self.assertEqual(tenant_user(self.email)["jobTitle"], "Analyst")

	def test_exit_offboards_and_return_enables(self):
		"""Left → licenses and groups removed, sign-in blocked, signed out; Active again → restored."""
		self.employee.reload()
		self.employee.status = "Left"
		self.employee.relieving_date = "2026-10-01"
		self.employee.save()
		user = tenant_user(self.email)
		self.assertFalse(user["accountEnabled"])
		self.assertEqual(user["sessionsRevoked"], 1)
		self.assertEqual(tenant_licenses(self.email), set())
		self.assertEqual(tenant_groups(self.email), set())
		account = frappe.get_doc("Microsoft 365 Account", self.employee.name)
		self.assertEqual(
			(account.status, len(account.licenses), len(account.groups)), ("Disabled", 0, 0)
		)
		mails = len(test_mode.load()["mail"])
		self.employee.reload()
		self.employee.status = "Active"
		self.employee.relieving_date = None
		self.employee.save()
		self.assertTrue(tenant_user(self.email)["accountEnabled"])
		self.assertEqual(tenant_licenses(self.email), {E3})
		self.assertEqual(
			frappe.db.get_value("Microsoft 365 Account", self.employee.name, "status"), "Active"
		)
		self.assertEqual(len(test_mode.load()["mail"]), mails + 1)

	def test_disabled_accounts_deleted_after_days(self):
		"""Delete After (Days) removes long-disabled users from Microsoft 365."""
		lifecycle.disable_account(self.employee.name)
		frappe.db.set_value(
			"Microsoft 365 Account", self.employee.name, "disabled_on", add_days(now_datetime(), -40)
		)
		frappe.db.set_single_value("Outlook Settings", "delete_after_days", 30)
		self.assertEqual(lifecycle.delete_expired(), 1)
		self.assertEqual(
			frappe.db.get_value("Microsoft 365 Account", self.employee.name, "status"), "Deleted"
		)
		self.assertRaises(Exception, tenant_user, self.email)

	def test_reset_password_emails_employee(self):
		"""Reset Password sends a new temporary password to the personal email."""
		self.assertEqual(lifecycle.reset_password(self.employee.name), [])
		self.assertEqual(test_mode.load()["mail"][-1]["to"], ["hana.personal@gmail.test"])

# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Onboarding: addresses, account creation, rules, manager, welcome email."""

from unittest.mock import patch

import frappe
from frappe.tests import IntegrationTestCase

from ms_integration.ms_core.services import test_mode
from ms_integration.ms_core.services.graph import GraphClient
from ms_integration.ms_outlook.services.addresses import last_name_base
from ms_integration.tests.m365_fixtures import (
	BUSINESS,
	DOMAIN,
	E3,
	PASSWORD,
	make_employee,
	make_rule,
	setup_microsoft365,
	tenant_groups,
	tenant_licenses,
	tenant_user,
)


class TestOnboarding(IntegrationTestCase):
	"""New employees get a Microsoft 365 account built from the settings and rules."""

	def setUp(self):
		"""
		Fresh Test Mode tenant and settings.

		Returns:
		        None
		"""
		setup_microsoft365()

	def tearDown(self):
		"""
		Stop patches.

		Returns:
		        None
		"""
		patch.stopall()

	def test_prefixes_dropped_only_as_separate_words(self):
		"""Al Harbi / Al-Harbi → harbi; Alharbi, Ellis and accents handled."""
		self.assertEqual(last_name_base("Al Harbi", ["al", "el"]), "harbi")
		self.assertEqual(last_name_base("al-Harbi", ["al", "el"]), "harbi")
		self.assertEqual(last_name_base("Alharbi", ["al", "el"]), "alharbi")
		self.assertEqual(last_name_base("Ellis", ["al", "el"]), "ellis")
		self.assertEqual(last_name_base("Núñez", []), "nunez")

	def test_active_employee_gets_account_licenses_groups_and_welcome(self):
		"""Matching rules add up; email saved on the Employee; welcome mail sent; no password stored."""
		make_rule("Everyone", licenses=[E3], groups=["All Staff"])
		make_rule("Egypt", 'doc.branch == "Egypt"', groups=["Egypt Office"])
		frappe.get_doc({"doctype": "Branch", "branch": "Egypt"}).insert(ignore_if_duplicate=True)
		employee = make_employee("Omar", "Al Saud", branch="Egypt", designation=None)
		account = frappe.get_doc("Microsoft 365 Account", employee.name)
		email = f"omar.s@{DOMAIN}"
		self.assertEqual((account.status, account.user_principal_name), ("Active", email))
		self.assertEqual(frappe.db.get_value("Employee", employee.name, "company_email"), email)
		self.assertEqual(tenant_licenses(email), {E3})
		self.assertEqual(tenant_groups(email), {"All Staff", "Egypt Office"})
		self.assertEqual({r.license for r in account.licenses}, {E3})
		self.assertEqual(len(account.groups), 2)
		user = tenant_user(email)
		self.assertEqual((user["usageLocation"], user["employeeId"]), ("SA", employee.name))
		mail = test_mode.load()["mail"]
		self.assertEqual(mail[-1]["to"], ["omar.personal@gmail.test"])
		stored = frappe.as_json(frappe.get_all("Microsoft 365 Log", fields=["details", "error"]))
		self.assertNotIn(PASSWORD, stored + frappe.as_json(account.as_dict()))

	def test_address_grows_until_free_in_microsoft_and_erpnext(self):
		"""john.s taken in Microsoft 365, john.sm taken in ERPNext → john.smi."""
		test_mode.new_user(tenant := test_mode.load(), {"userPrincipalName": f"john.s@{DOMAIN}"})
		test_mode.save(tenant)
		frappe.get_doc({"doctype": "User", "email": f"john.sm@{DOMAIN}", "first_name": "J"}).insert()
		employee = make_employee("John", "Smith")
		self.assertEqual(
			frappe.db.get_value("Microsoft 365 Account", employee.name, "user_principal_name"),
			f"john.smi@{DOMAIN}",
		)

	def test_no_free_address_fails_then_manual_address_works(self):
		"""Whole last name taken → Failed with a reason; typed address + retry → Active."""
		tenant = test_mode.load()
		for local in ("li.w", "li.wu"):
			test_mode.new_user(tenant, {"userPrincipalName": f"{local}@{DOMAIN}"})
		test_mode.save(tenant)
		employee = make_employee("Li", "Wu")
		account = frappe.get_doc("Microsoft 365 Account", employee.name)
		self.assertEqual(account.status, "Failed")
		self.assertIn("No free email address", account.error)
		account.user_principal_name = f"li.wu2@{DOMAIN}"
		account.save()
		frappe.get_attr("ms_integration.ms_outlook.tasks.provision_account")(employee.name)
		self.assertEqual(frappe.db.get_value("Microsoft 365 Account", employee.name, "status"), "Active")
		self.assertTrue(tenant_user(f"li.wu2@{DOMAIN}"))

	def test_missing_required_field_then_fixed(self):
		"""An empty required field fails the account; saving the fix creates it."""
		frappe.db.set_single_value("Outlook Settings", "required_fields", "first_name\ncell_number")
		employee = make_employee("Sara", "Khan")
		account = frappe.get_doc("Microsoft 365 Account", employee.name)
		self.assertEqual(account.status, "Failed")
		self.assertIn("Mobile", account.error)
		employee.reload()
		employee.cell_number = "+966500000000"
		employee.save()
		self.assertEqual(frappe.db.get_value("Microsoft 365 Account", employee.name, "status"), "Active")
		self.assertEqual(tenant_user(f"sara.k@{DOMAIN}")["mobilePhone"], "+966500000000")

	def test_not_created_before_condition_matches(self):
		"""An Inactive employee gets no account."""
		employee = make_employee("Ali", "Raza", status="Inactive")
		self.assertFalse(frappe.db.exists("Microsoft 365 Account", employee.name))

	def test_manager_and_erpnext_user(self):
		"""Manager set when the manager has an account; ERPNext User created and linked."""
		frappe.db.set_single_value("Outlook Settings", "create_erpnext_user", 1)
		manager = make_employee("Maya", "Iyer")
		employee = make_employee("Ravi", "Das", reports_to=manager.name)
		manager_id = frappe.db.get_value("Microsoft 365 Account", manager.name, "object_id")
		self.assertEqual(tenant_user(f"ravi.d@{DOMAIN}")["manager"], manager_id)
		self.assertEqual(frappe.db.get_value("Employee", employee.name, "user_id"), f"ravi.d@{DOMAIN}")

	def test_license_failure_is_reported_not_fatal(self):
		"""No usage location: the account is created, the license failure is shown and logged."""
		frappe.db.set_single_value("Outlook Settings", "usage_location", None)
		make_rule("Everyone", licenses=[BUSINESS])
		employee = make_employee("Noor", "Ali")
		account = frappe.get_doc("Microsoft 365 Account", employee.name)
		self.assertEqual(account.status, "Active")
		self.assertIn("usage location", account.error)
		self.assertTrue(
			frappe.db.exists("Microsoft 365 Log", {"action": "Licenses", "status": "Failed"})
		)
		self.assertEqual(
			GraphClient().get("/organization")["value"][0]["displayName"], "Test Mode Tenant"
		)

# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Shared test setup: Test Mode tenant, inline jobs, employees and rules."""

from unittest.mock import patch

import frappe

from ms_integration.ms_core.services import test_mode
from ms_integration.ms_core.services.graph import GraphClient
from ms_integration.ms_outlook.services.directory import sync_groups, sync_licenses

COMPANY = "M365 Test Company"
DOMAIN = "example.com"
PASSWORD = "Known#Pass1234567"
E3 = "ENTERPRISEPACK"
BUSINESS = "O365_BUSINESS_PREMIUM"
EXCHANGE = "EXCHANGESTANDARD"
JOB_OPTIONS = ("queue", "enqueue_after_commit", "now", "timeout", "job_id", "deduplicate")


def setup_microsoft365(**values) -> None:
	"""
	Enable the connection (Test Mode) and Outlook with a fresh fake tenant, clean records,
	synced licenses and groups, and jobs running inline.

	Parameters:
	        **values: Settings to override.

	Returns:
	        None
	"""
	for doctype in ("Microsoft 365 Account", "Microsoft 365 Provisioning Rule", "Microsoft 365 Log"):
		frappe.db.delete(doctype)
	frappe.db.delete("Employee", {"company": COMPANY})
	frappe.db.delete("User", {"name": ("like", f"%@{DOMAIN}")})
	connection = frappe.get_single("Microsoft 365 Settings")
	connection.update({"enabled": 1, "test_mode": 1})
	connection.save()
	settings = frappe.get_single("Outlook Settings")
	settings.update(
		{
			"enabled": 1,
			"email_domain": DOMAIN,
			"email_pattern": "{first}.{last}",
			"min_last_letters": 1,
			"drop_prefixes": "al\nel",
			"usage_location": "SA",
			"auto_create": 1,
			"create_condition": 'doc.status == "Active"',
			"disable_condition": 'doc.status in ("Inactive", "Left")',
			"required_fields": "first_name\nlast_name",
			"email_field": "company_email",
			"create_erpnext_user": 0,
			"manager_field": "reports_to",
			"sync_updates": 1,
			"remove_licenses_on_exit": 1,
			"remove_groups_on_exit": 1,
			"revoke_sessions": 1,
			"enable_on_return": 1,
			"delete_after_days": 0,
			"send_welcome_email": 1,
			"welcome_recipient_field": "personal_email",
			"welcome_via": "Microsoft 365",
			"sender_mailbox": f"it.admin@{DOMAIN}",
			**values,
		}
	)
	settings.save()
	test_mode.reset()
	client = GraphClient()
	sync_licenses(client)
	sync_groups(client)
	patch.object(frappe.db, "commit", lambda *a, **k: None).start()
	patch(
		"ms_integration.ms_outlook.services.provisioning.new_password",
		return_value=PASSWORD,
	).start()
	run_jobs_inline()


def run_jobs_inline():
	"""
	Run frappe.enqueue jobs immediately, dropping enqueue options.

	Returns:
	        MagicMock: The patched enqueue.
	"""

	def run_now(method, **kwargs):
		"""
		Call the job directly.

		Parameters:
		        method (str, required): Dotted path.
		        **kwargs: Job arguments plus enqueue options.

		Returns:
		        Any: Job result.
		"""
		for option in JOB_OPTIONS:
			kwargs.pop(option, None)
		return frappe.get_attr(method)(**kwargs)

	return patch("frappe.enqueue", side_effect=run_now).start()


def group(name: str) -> str:
	"""
	Microsoft 365 Group record name by display name.

	Parameters:
	        name (str, required): Group display name.

	Returns:
	        str: Group name (object id).
	"""
	return frappe.db.get_value("Microsoft 365 Group", {"display_name": name})


def make_rule(name: str, condition: str = "", licenses=(), groups=()) -> None:
	"""
	A provisioning rule.

	Parameters:
	        name (str, required): Rule name.
	        condition (str, optional): Python condition.
	        licenses (Iterable[str], optional): License names.
	        groups (Iterable[str], optional): Group display names.

	Returns:
	        None
	"""
	frappe.get_doc(
		{
			"doctype": "Microsoft 365 Provisioning Rule",
			"rule_name": name,
			"condition": condition,
			"licenses": [{"license": lic} for lic in licenses],
			"groups": [{"group": group(g)} for g in groups],
		}
	).insert()


def ensure_company() -> str:
	"""
	A company for test employees.

	Returns:
	        str: Company name.
	"""
	if not frappe.db.exists("Warehouse Type", "Transit"):
		frappe.get_doc({"doctype": "Warehouse Type", "name": "Transit"}).insert()
	if not frappe.db.exists("Company", COMPANY):
		frappe.get_doc(
			{
				"doctype": "Company",
				"company_name": COMPANY,
				"abbr": "M3T",
				"default_currency": "EGP",
				"country": "Egypt",
			}
		).insert()
	for gender in ("Male", "Female"):
		if not frappe.db.exists("Gender", gender):
			frappe.get_doc({"doctype": "Gender", "gender": gender}).insert()
	return COMPANY


def make_employee(first: str, last: str, **values):
	"""
	An Employee (Active unless given), saved with the hooks running.

	Parameters:
	        first (str, required): First name.
	        last (str, required): Last name.
	        **values: Other Employee fields.

	Returns:
	        Document: The Employee.
	"""
	return frappe.get_doc(
		{
			"doctype": "Employee",
			"first_name": first,
			"last_name": last,
			"gender": "Male",
			"date_of_birth": "1990-01-01",
			"date_of_joining": "2026-01-01",
			"company": ensure_company(),
			"status": "Active",
			"personal_email": f"{first.lower()}.personal@gmail.test",
			**values,
		}
	).insert()


def tenant_user(email: str) -> dict:
	"""
	A user of the fake tenant by sign-in name.

	Parameters:
	        email (str, required): Sign-in name.

	Returns:
	        dict: Stored user.
	"""
	return test_mode.find_user(test_mode.load(), email)


def tenant_licenses(email: str) -> set[str]:
	"""
	SKU part numbers assigned to a fake-tenant user.

	Parameters:
	        email (str, required): Sign-in name.

	Returns:
	        set[str]: Part numbers.
	"""
	tenant = test_mode.load()
	return {tenant["skus"][sku]["skuPartNumber"] for sku in tenant_user(email)["licenses"]}


def tenant_groups(email: str) -> set[str]:
	"""
	Display names of the groups a fake-tenant user belongs to.

	Parameters:
	        email (str, required): Sign-in name.

	Returns:
	        set[str]: Group names.
	"""
	tenant = test_mode.load()
	user_id = tenant_user(email)["id"]
	return {g["displayName"] for g in tenant["groups"].values() if user_id in g["members"]}

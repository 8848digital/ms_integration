# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Decide what an Employee change means for the Microsoft 365 account."""

import frappe

from ms_integration.ms_core.services.settings import matches
from ms_integration.ms_outlook.services.payload import profile
from ms_integration.ms_outlook.services.settings import get_settings, is_enabled
from ms_integration.ms_outlook.tasks import enqueue

NOT_CREATED = (None, "Pending", "Failed", "Deleted")


def queue_account_changes(employee) -> None:
	"""
	Queue the right job for the employee's account:
	disable on exit, create or enable when the create condition matches,
	otherwise update changed details.

	Parameters:
	        employee (Document, required): Saved Employee.

	Returns:
	        None
	"""
	settings = get_settings()
	if not is_enabled():
		return
	status = frappe.db.get_value("Microsoft 365 Account", employee.name, "status")
	if (settings.disable_condition or "").strip() and matches(settings.disable_condition, employee):
		if status == "Active":
			enqueue("disable_account", employee.name, account=employee.name)
		return
	if status == "Active":
		if settings.sync_updates and details_changed(employee, settings):
			enqueue("update_account", employee.name, account=employee.name)
		return
	if not matches(settings.create_condition, employee):
		return
	if status == "Disabled" and settings.enable_on_return:
		enqueue("enable_account", employee.name, account=employee.name)
	elif status in NOT_CREATED and settings.auto_create:
		enqueue("provision_account", employee.name, employee=employee.name)


def details_changed(employee, settings) -> bool:
	"""
	Whether a property sent to Microsoft 365 (or the manager) changed in this save.

	Parameters:
	        employee (Document, required): Saved Employee.
	        settings (Document, required): Microsoft 365 Settings.

	Returns:
	        bool: True when Microsoft 365 needs an update.
	"""
	before = employee.get_doc_before_save()
	if not before:
		return False
	if settings.manager_field and before.get(settings.manager_field) != employee.get(
		settings.manager_field
	):
		return True
	return profile(before, settings) != profile(employee, settings)

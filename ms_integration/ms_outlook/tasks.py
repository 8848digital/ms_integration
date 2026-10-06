# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Background and scheduled Microsoft 365 jobs."""

import frappe

from ms_integration.ms_core.services.graph import GraphClient
from ms_integration.ms_core.services.logs import log
from ms_integration.ms_outlook.services import (
	directory,
	document_groups,
	groups_api,
	lifecycle,
	provisioning,
)
from ms_integration.ms_outlook.services.settings import get_settings, is_enabled

TASKS = "ms_integration.ms_outlook.tasks"


def enqueue(task: str, key: str, **kwargs) -> None:
	"""
	Queue one of this module's jobs after the current transaction commits;
	the same job for the same key is not queued twice.

	Parameters:
	        task (str, required): Function name in this module.
	        key (str, required): Employee / account name used for de-duplication.
	        **kwargs: Arguments for the job.

	Returns:
	        None
	"""
	frappe.enqueue(
		f"{TASKS}.{task}",
		queue="long" if task == "sync_directory" else "default",
		job_id=f"microsoft365:{task}:{key}",
		deduplicate=True,
		enqueue_after_commit=True,
		**kwargs,
	)


def provision_account(employee: str) -> None:
	"""
	Create the employee's Microsoft 365 account.

	Parameters:
	        employee (str, required): Employee name.

	Returns:
	        None
	"""
	if is_enabled():
		provisioning.provision(employee)


def update_account(account: str) -> None:
	"""
	Send changed employee details to Microsoft 365.

	Parameters:
	        account (str, required): Microsoft 365 Account name.

	Returns:
	        None
	"""
	if is_enabled():
		lifecycle.update_account(account)


def disable_account(account: str) -> None:
	"""
	Offboard the account.

	Parameters:
	        account (str, required): Microsoft 365 Account name.

	Returns:
	        None
	"""
	if is_enabled():
		lifecycle.disable_account(account)


def enable_account(account: str) -> None:
	"""
	Enable the account again (rehire).

	Parameters:
	        account (str, required): Microsoft 365 Account name.

	Returns:
	        None
	"""
	if is_enabled():
		lifecycle.enable_account(account)


def sync_directory() -> None:
	"""
	Licenses, groups and accounts from Microsoft 365 (button or daily).

	Returns:
	        None
	"""
	if not is_enabled():
		return
	try:
		directory.sync_all()
	except Exception as e:
		log("Directory Sync", False, error=e)


def daily_sync() -> None:
	"""
	Scheduled: directory sync when switched on in settings.

	Returns:
	        None
	"""
	if is_enabled() and get_settings().sync_directory_daily:
		sync_directory()
		resync_document_groups()


def sync_document_group(doctype: str, name: str) -> None:
	"""
	Create / update / close the document's Microsoft 365 group.

	Parameters:
	        doctype (str, required): DocType.
	        name (str, required): Document name.

	Returns:
	        None
	"""
	if is_enabled():
		document_groups.sync_document(doctype, name)


def delete_group(group_id: str) -> None:
	"""
	Delete a group whose document was deleted.

	Parameters:
	        group_id (str, required): Group object id.

	Returns:
	        None
	"""
	if not is_enabled():
		return
	try:
		groups_api.delete_group(GraphClient(), group_id)
		log("Document Group", details=f"Deleted group {group_id}")
	except Exception as e:
		log("Document Group", False, details=f"Delete group {group_id}", error=e)


def resync_document_groups() -> None:
	"""
	Re-check members of every active document group (catches changes made
	without a save, e.g. imports).

	Returns:
	        None
	"""
	rows = frappe.get_all(
		"Microsoft 365 Document Group", {"status": "Active"}, ["reference_doctype", "reference_name"]
	)
	for row in rows:
		document_groups.sync_document(row.reference_doctype, row.reference_name)


def delete_expired_accounts() -> None:
	"""
	Scheduled: delete accounts disabled longer than the configured days.

	Returns:
	        None
	"""
	if is_enabled():
		lifecycle.delete_expired()

# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Microsoft 365 Log entries: every change made in Microsoft 365 and every failure."""

import frappe

from ms_integration.ms_core.services.graph import friendly


def log(
	action: str,
	ok: bool = True,
	account: str | None = None,
	details: str | None = None,
	error: Exception | str | None = None,
) -> None:
	"""
	Write a Microsoft 365 Log (never raises; passwords are never passed here).

	Parameters:
	        action (str, required): One of the Log's Action options.
	        ok (bool, optional): Success or Failed.
	        account (str, optional): Microsoft 365 Account name (the Employee).
	        details (str, optional): What changed.
	        error (Exception | str, optional): Failure.

	Returns:
	        None
	"""
	try:
		frappe.get_doc(
			{
				"doctype": "Microsoft 365 Log",
				"action": action,
				"status": "Success" if ok else "Failed",
				"account": account if account and frappe.db.exists("Microsoft 365 Account", account) else None,
				"reference_doctype": "Employee" if account else None,
				"reference_name": account,
				"details": details,
				"error": friendly(error) if isinstance(error, Exception) else error,
			}
		).insert(ignore_permissions=True)
	except Exception:
		frappe.log_error("Microsoft 365 | could not write log")


def summary(added=(), removed=(), label: str = "") -> str:
	"""
	One-line "added / removed" summary for a log.

	Parameters:
	        added (Iterable[str], optional): Names added.
	        removed (Iterable[str], optional): Names removed.
	        label (str, optional): What the names are (licenses, groups).

	Returns:
	        str: Summary text.
	"""
	parts = []
	if added:
		parts.append(f"Added {label}: {', '.join(added)}")
	if removed:
		parts.append(f"Removed {label}: {', '.join(removed)}")
	return "\n".join(parts) or "No change"

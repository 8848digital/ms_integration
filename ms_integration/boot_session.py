# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Data added to frappe.boot for this app."""

import frappe


def boot_session(bootinfo) -> None:
	"""
	DocTypes that have a Microsoft 365 Group Rule (for the form button).

	Parameters:
	        bootinfo (dict, required): Boot data being built.

	Returns:
	        None
	"""
	if frappe.session.user == "Guest":
		return
	from ms_integration.ms_outlook.services.group_rules import rule_doctypes

	bootinfo.ms_integration = {"group_doctypes": rule_doctypes()["targets"]}

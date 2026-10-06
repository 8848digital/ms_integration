# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Outlook Settings access."""

import frappe

from ms_integration.ms_core.services.settings import is_connected

SETTINGS = "Outlook Settings"


def get_settings():
	"""
	The cached Outlook Settings (accounts, offboarding, welcome email, sync).

	Returns:
	        Document: Outlook Settings.
	"""
	return frappe.get_cached_doc(SETTINGS)


def is_enabled() -> bool:
	"""
	Whether the Outlook features run: the Microsoft 365 connection and Outlook
	Settings are both enabled.

	Returns:
	        bool: True when both are switched on.
	"""
	return is_connected() and bool(get_settings().enabled)

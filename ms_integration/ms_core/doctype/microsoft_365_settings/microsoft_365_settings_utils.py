# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Save logic for Microsoft 365 Settings."""

import frappe

from ms_integration.ms_core.services.graph import token_cache_key


def after_save(doc) -> None:
	"""
	Forget the cached access token (credentials may have changed).

	Parameters:
	        doc (Document, required): Microsoft 365 Settings.

	Returns:
	        None
	"""
	before = doc.get_doc_before_save()
	if before:
		frappe.cache.delete_value(token_cache_key(before))
	frappe.cache.delete_value(token_cache_key(doc))

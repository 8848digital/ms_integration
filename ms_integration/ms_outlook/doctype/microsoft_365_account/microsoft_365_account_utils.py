# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Validation and Microsoft 365 push for Microsoft 365 Account."""

from ms_integration.ms_outlook.services.access import push_changes


def validate_account(doc) -> None:
	"""
	Normalise the sign-in name and push row changes of an active account.

	Parameters:
	        doc (Document, required): Microsoft 365 Account.

	Returns:
	        None
	"""
	doc.user_principal_name = (doc.user_principal_name or "").strip().lower() or None
	push_changes(doc)

# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Validation and cache refresh for Microsoft 365 Group Rule."""

from ms_integration.ms_outlook.services.group_rules import clear_cache, validate_rule


def validate_group_rule(doc) -> None:
	"""
	Validate the rule.

	Parameters:
	        doc (Document, required): Microsoft 365 Group Rule.

	Returns:
	        None
	"""
	validate_rule(doc)


def after_change(doc) -> None:
	"""
	Forget the cached rule map after a save or delete.

	Parameters:
	        doc (Document, required): Microsoft 365 Group Rule.

	Returns:
	        None
	"""
	clear_cache()

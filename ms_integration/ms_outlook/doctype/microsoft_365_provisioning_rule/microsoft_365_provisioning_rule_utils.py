# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Validation for Microsoft 365 Provisioning Rule."""

from frappe import _

from ms_integration.ms_core.services.settings import validate_condition
from ms_integration.ms_outlook.services.access import check_groups_manageable


def validate_rule(doc) -> None:
	"""
	Valid condition; only groups whose members can be changed through Graph.

	Parameters:
	        doc (Document, required): Microsoft 365 Provisioning Rule.

	Returns:
	        None
	"""
	validate_condition(doc.condition, _("Condition"))
	check_groups_manageable({row.group for row in doc.groups})

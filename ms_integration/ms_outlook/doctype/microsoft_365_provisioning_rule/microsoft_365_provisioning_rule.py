# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Microsoft 365 Provisioning Rule: licenses and groups for matching employees."""

from frappe.model.document import Document

from ms_integration.ms_outlook.doctype.microsoft_365_provisioning_rule.microsoft_365_provisioning_rule_utils import (
	validate_rule,
)


class Microsoft365ProvisioningRule(Document):
	"""Condition on the Employee → licenses and groups to give."""

	def validate(self):
		"""
		Check the condition and that every group's members can be changed.

		Returns:
		        None
		"""
		validate_rule(self)

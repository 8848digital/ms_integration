# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Microsoft 365 Group Rule: a Microsoft 365 group per document of one DocType."""

from frappe.model.document import Document

from ms_integration.ms_outlook.doctype.microsoft_365_group_rule.microsoft_365_group_rule_utils import (
	after_change,
	validate_group_rule,
)


class Microsoft365GroupRule(Document):
	"""When to create the group, who is in it, and what happens on close."""

	def validate(self):
		"""
		Check conditions, field paths, sources and write-back fields.

		Returns:
		        None
		"""
		validate_group_rule(self)

	def on_update(self):
		"""
		Refresh the cached list of DocTypes with rules.

		Returns:
		        None
		"""
		after_change(self)

	def on_trash(self):
		"""
		Refresh the cached list of DocTypes with rules.

		Returns:
		        None
		"""
		after_change(self)

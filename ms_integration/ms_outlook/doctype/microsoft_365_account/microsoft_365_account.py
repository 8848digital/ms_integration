# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Microsoft 365 Account: one employee's account, licenses and groups."""

from frappe.model.document import Document

from ms_integration.ms_outlook.doctype.microsoft_365_account.microsoft_365_account_utils import (
	validate_account,
)


class Microsoft365Account(Document):
	"""Mirror of the employee's Microsoft 365 user; editing rows changes Microsoft 365."""

	def validate(self):
		"""
		Push added/removed licenses and groups to Microsoft 365.

		Returns:
		        None
		"""
		validate_account(self)

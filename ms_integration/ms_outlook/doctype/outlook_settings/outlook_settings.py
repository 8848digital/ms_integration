# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Outlook Settings: accounts, offboarding, welcome email and directory sync."""

from frappe.model.document import Document

from ms_integration.ms_outlook.doctype.outlook_settings.outlook_settings_utils import (
	validate_outlook_settings,
)


class OutlookSettings(Document):
	"""Single DocType holding the Outlook (accounts, licenses, groups) options."""

	def validate(self):
		"""
		Check conditions, domain, pattern and usage location.

		Returns:
		        None
		"""
		validate_outlook_settings(self)

# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Microsoft 365 Settings: the shared Microsoft 365 connection (tenant, app, Test Mode)."""

from frappe.model.document import Document

from ms_integration.ms_core.doctype.microsoft_365_settings.microsoft_365_settings_utils import (
	after_save,
)


class Microsoft365Settings(Document):
	"""Single DocType holding the Microsoft 365 connection used by every module."""

	def on_update(self):
		"""
		Drop the cached token so changed credentials are used at once.

		Returns:
		        None
		"""
		after_save(self)

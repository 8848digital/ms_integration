# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""doc_events for Employee (owned by ERPNext)."""

from ms_integration.ms_outlook.customization.employee.employee_utils import queue_account_changes


def on_update(doc, method=None):
	"""
	After an Employee is saved: create, update, disable or enable the
	Microsoft 365 account as the settings say.

	Parameters:
	        doc (Document, required): Employee.
	        method (str, optional): Hook name.

	Returns:
	        None
	"""
	queue_account_changes(doc)

# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

app_name = "ms_integration"
app_title = "MS Integration"
app_publisher = "8848 Digital LLP"
app_description = (
	"Microsoft 365 integrations for Frappe/ERPNext: Outlook (accounts, licenses, groups)"
)
app_email = "mahak@8848digital.com"
app_license = "Proprietary"

required_apps = ["erpnext"]

custom_fixtures = [
	{"dt": "Custom Field", "filters": {"module": ("in", ["MS Core", "MS Outlook"])}}
]

commands = ["ms_integration.commands.export_fixtures.export_fixtures"]

after_request = [
	"ms_integration.utils.api_handlers.response_formatter.format_frappe_response_to_custom"
]

after_install = "ms_integration.install.after_install"

app_include_js = "ms_integration.bundle.js"

boot_session = "ms_integration.boot_session.boot_session"

doctype_js = {"Employee": "ms_outlook/customization/employee/employee.js"}

GROUP_EVENTS = "ms_integration.ms_outlook.group_events"

doc_events = {
	# Group Rules can target any DocType; exits after a cached lookup otherwise.
	"*": {
		"on_update": f"{GROUP_EVENTS}.on_update",
		"on_update_after_submit": f"{GROUP_EVENTS}.on_update",
		"on_trash": f"{GROUP_EVENTS}.on_trash",
		"after_delete": f"{GROUP_EVENTS}.after_delete",
	},
	"Employee": {
		"on_update": "ms_integration.ms_outlook.customization.employee.employee.on_update",
	},
}

TASKS = "ms_integration.ms_outlook.tasks"

scheduler_events = {
	"daily_long": [f"{TASKS}.daily_sync"],
	"daily": [f"{TASKS}.delete_expired_accounts"],
}

// Copyright (c) 2026 8848 Digital LLP. All rights reserved.
// Proprietary and confidential. Unauthorized copying, distribution, or use
// of this file, via any medium, is strictly prohibited without prior
// written permission from 8848 Digital LLP.

frappe.ui.form.on("Microsoft 365 Settings", {
	refresh(frm) {
		frm.add_custom_button(__("Test Connection"), () =>
			ms_integration.call("ms_core.api.v1.settings.test_connection")
		);
		frm.add_custom_button(
			__("Outlook Settings"),
			() => frappe.set_route("Form", "Outlook Settings"),
			__("Go to")
		);
		frm.add_custom_button(
			__("Log"),
			() => frappe.set_route("List", "Microsoft 365 Log"),
			__("Go to")
		);
	},
});

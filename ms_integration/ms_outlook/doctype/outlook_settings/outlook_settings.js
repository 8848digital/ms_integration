// Copyright (c) 2026 8848 Digital LLP. All rights reserved.
// Proprietary and confidential. Unauthorized copying, distribution, or use
// of this file, via any medium, is strictly prohibited without prior
// written permission from 8848 Digital LLP.

frappe.ui.form.on("Outlook Settings", {
	refresh(frm) {
		frm.add_custom_button(__("Sync Licenses, Groups and Accounts"), () =>
			ms_integration.call("ms_outlook.api.v1.directory.sync_directory")
		);
		const go = (label, route) =>
			frm.add_custom_button(label, () => frappe.set_route(...route), __("Go to"));
		go(__("Connection"), ["Form", "Microsoft 365 Settings"]);
		go(__("Accounts"), ["List", "Microsoft 365 Account"]);
		go(__("Provisioning Rules"), ["List", "Microsoft 365 Provisioning Rule"]);
		go(__("Group Rules"), ["List", "Microsoft 365 Group Rule"]);
		go(__("Log"), ["List", "Microsoft 365 Log"]);
	},
});

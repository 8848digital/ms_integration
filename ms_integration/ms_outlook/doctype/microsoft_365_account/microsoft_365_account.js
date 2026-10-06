// Copyright (c) 2026 8848 Digital LLP. All rights reserved.
// Proprietary and confidential. Unauthorized copying, distribution, or use
// of this file, via any medium, is strictly prohibited without prior
// written permission from 8848 Digital LLP.

const STATUS_COLORS = {
	Pending: "gray",
	Creating: "blue",
	Active: "green",
	Disabled: "orange",
	Deleted: "darkgrey",
	Failed: "red",
};

frappe.ui.form.on("Microsoft 365 Account", {
	setup(frm) {
		const manageable = { filters: { membership_managed: 1, removed: 0 } };
		frm.set_query("group", "groups", () => manageable);
	},

	refresh(frm) {
		const status = frm.doc.status;
		frm.page.set_indicator(__(status), STATUS_COLORS[status] || "gray");
		if (frm.is_new()) return;
		const run = (action, label, confirmText) => {
			const go = () =>
				ms_integration.call(
					"ms_outlook.api.v1.account.run_action",
					{ account: frm.doc.name, action },
					() => setTimeout(() => frm.reload_doc(), 1500)
				);
			frm.add_custom_button(
				label,
				() => (confirmText ? frappe.confirm(confirmText, go) : go()),
				__("Actions")
			);
		};
		if (["Pending", "Failed", "Deleted"].includes(status)) {
			frm.add_custom_button(status === "Pending" ? __("Create Account") : __("Retry"), () =>
				ms_integration.call(
					"ms_outlook.api.v1.account.create_account",
					{ employee: frm.doc.employee },
					() => setTimeout(() => frm.reload_doc(), 2000)
				)
			);
		}
		if (status === "Active") {
			run(
				"reset_password",
				__("Reset Password and Email It"),
				__("Set a new temporary password and email it to the employee?")
			);
			run(
				"disable",
				__("Disable (Offboard)"),
				__("Remove licenses and groups and block sign-in?")
			);
		}
		if (status === "Disabled") run("enable", __("Enable Again"));
		if (frm.doc.object_id) run("refresh", __("Refresh from Microsoft 365"));
		frm.toggle_enable(["licenses", "groups"], status === "Active");
	},
});

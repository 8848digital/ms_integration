// Copyright (c) 2026 8848 Digital LLP. All rights reserved.
// Proprietary and confidential. Unauthorized copying, distribution, or use
// of this file, via any medium, is strictly prohibited without prior
// written permission from 8848 Digital LLP.

// Employee form: Microsoft 365 account status and actions.

frappe.ui.form.on("Employee", {
	refresh(frm) {
		if (frm.is_new()) return;
		ms_integration.call(
			"ms_outlook.api.v1.account.get_status",
			{ employee: frm.doc.name },
			(data) => showMicrosoft365(frm, data),
			"GET"
		);
	},
});

/**
 * Headline with the account status, and Open / Create buttons.
 *
 * @param {object} frm - Employee form.
 * @param {object} data - {enabled, account, status, email, error}.
 * @returns {void}
 */
function showMicrosoft365(frm, data) {
	if (!data.enabled) return;
	const group = __("Microsoft 365");
	if (data.account) {
		const color =
			{ Active: "green", Disabled: "orange", Failed: "red" }[data.status] || "blue";
		const text = data.email ? `${__(data.status)} · ${data.email}` : __(data.status);
		frm.dashboard.set_headline_alert(
			`<span class="indicator ${color}">${__("Microsoft 365")}: ${frappe.utils.escape_html(
				text
			)}</span>`
		);
		frm.add_custom_button(
			__("Open Account"),
			() => frappe.set_route("Form", "Microsoft 365 Account", data.account),
			group
		);
	}
	if (!data.account || ["Pending", "Failed", "Deleted"].includes(data.status)) {
		frm.add_custom_button(
			data.status === "Failed" ? __("Retry Account Creation") : __("Create Account"),
			() =>
				ms_integration.call(
					"ms_outlook.api.v1.account.create_account",
					{ employee: frm.doc.name },
					() => frm.reload_doc()
				),
			group
		);
	}
}

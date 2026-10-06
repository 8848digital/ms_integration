// Copyright (c) 2026 8848 Digital LLP. All rights reserved.
// Proprietary and confidential. Unauthorized copying, distribution, or use
// of this file, via any medium, is strictly prohibited without prior
// written permission from 8848 Digital LLP.

// Desk-wide helper for calling this app's APIs (standard response envelope).

frappe.provide("ms_integration");

/**
 * Call an API of this app; show its message, then run the callback.
 *
 * @param {string} method - Path after ms_integration., e.g. ms_outlook.api.v1.account.get_status.
 * @param {object} args - Arguments.
 * @param {Function} [done] - Called with the envelope's data on success.
 * @param {string} [type] - HTTP method (default POST).
 * @returns {void}
 */
ms_integration.call = function (method, args, done, type) {
	frappe.call({
		method: `ms_integration.${method}`,
		args: args || {},
		type: type || "POST",
		freeze: type !== "GET",
		callback: (r) => {
			if (type !== "GET" && r.message) {
				frappe.show_alert(
					{ message: r.message, indicator: r.status ? "green" : "red" },
					7
				);
			}
			if (r.status && done) done(r.data || {});
		},
		error: (xhr) => {
			const body = (xhr && xhr.responseJSON) || {};
			if (body.message) {
				frappe.msgprint({
					title: __("Microsoft 365"),
					message: body.message,
					indicator: "red",
				});
			}
		},
	});
};

$(document).on("form-refresh", (event, frm) => {
	const doctypes = (frappe.boot.ms_integration || {}).group_doctypes || [];
	if (frm.is_new() || !doctypes.includes(frm.doctype)) return;
	ms_integration.call(
		"ms_outlook.api.v1.group.get_group",
		{ doctype: frm.doctype, name: frm.doc.name },
		(data) => addGroupButtons(frm, data),
		"GET"
	);
});

/**
 * "Open Group" when the document has one, and "Create / Sync Group".
 *
 * @param {object} frm - Document form.
 * @param {object} data - {has_rule, record, status, display_name, mail}.
 * @returns {void}
 */
function addGroupButtons(frm, data) {
	if (!data.has_rule) return;
	const label = __("Microsoft 365");
	if (data.record) {
		frm.add_custom_button(
			__("Open Group"),
			() => frappe.set_route("Form", "Microsoft 365 Document Group", data.record),
			label
		);
	}
	frm.add_custom_button(
		data.record ? __("Sync Group Members") : __("Create Group"),
		() =>
			ms_integration.call("ms_outlook.api.v1.group.sync_group", {
				doctype: frm.doctype,
				name: frm.doc.name,
			}),
		label
	);
}

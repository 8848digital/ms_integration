// Copyright (c) 2026 8848 Digital LLP. All rights reserved.
// Proprietary and confidential. Unauthorized copying, distribution, or use
// of this file, via any medium, is strictly prohibited without prior
// written permission from 8848 Digital LLP.

frappe.ui.form.on("Microsoft 365 Provisioning Rule", {
	setup(frm) {
		// Only groups whose members can be changed through Microsoft Graph.
		frm.set_query("group", "groups", () => ({
			filters: { membership_managed: 1, removed: 0 },
		}));
	},
});

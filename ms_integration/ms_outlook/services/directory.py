# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""
Directory sync: licenses and groups of the tenant, each account's current
licenses and groups, and linking employees who already have an account.
"""

import frappe
from frappe.utils import now_datetime

from ms_integration.ms_core.services.graph import GraphClient
from ms_integration.ms_core.services.logs import log
from ms_integration.ms_outlook.services import directory_api as api
from ms_integration.ms_outlook.services.settings import get_settings

LICENSE = "Microsoft 365 License"
GROUP = "Microsoft 365 Group"
ACCOUNT = "Microsoft 365 Account"


def sync_all() -> dict:
	"""
	Full sync: licenses, groups, every account, then link existing accounts.

	Returns:
	        dict: Counts {licenses, groups, accounts, linked}.
	"""
	client = GraphClient()
	counts = {"licenses": sync_licenses(client), "groups": sync_groups(client)}
	counts["accounts"] = refresh_accounts(client)
	counts["linked"] = link_existing(client) if get_settings().link_existing_accounts else 0
	frappe.db.set_single_value("Microsoft 365 Settings", "last_directory_sync", now_datetime())
	log("Directory Sync", details=", ".join(f"{k}: {v}" for k, v in counts.items()))
	return counts


def sync_licenses(client: GraphClient) -> int:
	"""
	Upsert every license type the tenant has.

	Parameters:
	        client (GraphClient, required): Graph client.

	Returns:
	        int: Number of license types.
	"""
	skus = client.get_all("/subscribedSkus")
	for sku in skus:
		upsert_license(sku)
	return len(skus)


def upsert_license(sku: dict) -> str:
	"""
	Create or update one Microsoft 365 License.

	Parameters:
	        sku (dict, required): Graph subscribedSku (or licenseDetails row).

	Returns:
	        str: License name (SKU part number).
	"""
	name = sku["skuPartNumber"]
	doc = (
		frappe.get_doc(LICENSE, name) if frappe.db.exists(LICENSE, name) else frappe.new_doc(LICENSE)
	)
	doc.update({"sku_part_number": name, "sku_id": sku["skuId"], "label": doc.label or name})
	if "prepaidUnits" in sku:
		enabled = (sku.get("prepaidUnits") or {}).get("enabled") or 0
		used = sku.get("consumedUnits") or 0
		doc.update(
			{
				"enabled_units": enabled,
				"consumed_units": used,
				"available_units": max(enabled - used, 0),
				"capability_status": sku.get("capabilityStatus"),
				"last_synced": now_datetime(),
			}
		)
	doc.save(ignore_permissions=True)
	return doc.name


def sync_groups(client: GraphClient) -> int:
	"""
	Upsert every group; groups gone from the tenant are flagged as removed.

	Parameters:
	        client (GraphClient, required): Graph client.

	Returns:
	        int: Number of groups.
	"""
	groups = client.get_all("/groups", params={"$select": api.GROUP_FIELDS, "$top": "999"})
	seen = {upsert_group(group) for group in groups}
	for name in frappe.get_all(GROUP, {"removed": 0}, pluck="name"):
		if name not in seen:
			frappe.db.set_value(GROUP, name, "removed", 1)
	return len(groups)


def upsert_group(group: dict) -> str:
	"""
	Create or update one Microsoft 365 Group.

	Parameters:
	        group (dict, required): Graph group.

	Returns:
	        str: Group name (object id).
	"""
	name = group["id"]
	doc = frappe.get_doc(GROUP, name) if frappe.db.exists(GROUP, name) else frappe.new_doc(GROUP)
	kind = group_type(group)
	doc.update(
		{
			"group_id": name,
			"display_name": group.get("displayName"),
			"mail": group.get("mail"),
			"group_type": kind,
			"description": group.get("description"),
			"membership_managed": int(
				kind in ("Microsoft 365", "Security")
				and "DynamicMembership" not in (group.get("groupTypes") or [])
			),
			"removed": 0,
			"last_synced": now_datetime(),
		}
	)
	doc.save(ignore_permissions=True)
	return doc.name


def group_type(group: dict) -> str:
	"""
	Microsoft 365 / Security / Mail-enabled Security / Distribution.

	Parameters:
	        group (dict, required): Graph group.

	Returns:
	        str: Type label.
	"""
	if "Unified" in (group.get("groupTypes") or []):
		return "Microsoft 365"
	if group.get("securityEnabled"):
		return "Mail-enabled Security" if group.get("mailEnabled") else "Security"
	return "Distribution"


def refresh_accounts(client: GraphClient) -> int:
	"""
	Refresh licenses, groups and sign-in state of every active or disabled account.

	Parameters:
	        client (GraphClient, required): Graph client.

	Returns:
	        int: Accounts refreshed.
	"""
	count = 0
	for name in frappe.get_all(ACCOUNT, {"status": ("in", ("Active", "Disabled"))}, pluck="name"):
		try:
			refresh_account(client, frappe.get_doc(ACCOUNT, name))
			count += 1
		except Exception as e:
			log("Directory Sync", False, name, error=e)
	return count


def refresh_account(client: GraphClient, account, user: dict | None = None) -> None:
	"""
	Copy the account's current licenses, groups and sign-in state from Microsoft 365.

	Parameters:
	        client (GraphClient, required): Graph client.
	        account (Document, required): Microsoft 365 Account.
	        user (dict, optional): Graph user when already fetched.

	Returns:
	        None
	"""
	user = user or client.get(f"/users/{account.object_id}", params={"$select": api.USER_FIELDS})
	licenses = [
		upsert_license(row) for row in client.get_all(f"/users/{account.object_id}/licenseDetails")
	]
	groups = [upsert_group(group) for group in api.user_groups(client, account.object_id)]
	account.set("licenses", [{"license": name} for name in licenses])
	account.set("groups", [{"group": name} for name in groups])
	account.account_enabled = int(bool(user.get("accountEnabled")))
	account.user_principal_name = user.get("userPrincipalName") or account.user_principal_name
	account.last_synced = now_datetime()
	account.flags.from_sync = True
	account.save(ignore_permissions=True)


def link_existing(client: GraphClient) -> int:
	"""
	Create accounts for employees whose work email already exists in Microsoft 365.

	Parameters:
	        client (GraphClient, required): Graph client.

	Returns:
	        int: Accounts linked.
	"""
	field = get_settings().email_field
	if not field:
		return 0
	users = {}
	for user in api.all_users(client):
		for address in (user.get("userPrincipalName"), user.get("mail")):
			if address:
				users[address.lower()] = user
	linked = set(frappe.get_all(ACCOUNT, pluck="employee"))
	count = 0
	for employee, email in frappe.get_all(
		"Employee", {field: ("is", "set")}, ["name", field], as_list=1
	):
		user = users.get((email or "").lower())
		if user and employee not in linked:
			account = frappe.get_doc(
				{
					"doctype": ACCOUNT,
					"employee": employee,
					"object_id": user["id"],
					"status": "Active" if user.get("accountEnabled") else "Disabled",
				}
			)
			account.flags.from_sync = True
			refresh_account(client, account, user)
			count += 1
	return count

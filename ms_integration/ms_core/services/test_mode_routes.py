# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Graph routes answered by the Test Mode fake tenant."""

import re
import uuid

from ms_integration.ms_core.services.test_mode import (
	fail,
	find_user,
	group_reply,
	new_user,
	public,
)
from ms_integration.ms_core.services.test_mode_group_routes import (
	add_owner,
	create_group,
	delete_group,
	get_group,
	list_members,
	list_owners,
)


def organization(tenant, params, body):
	"""GET /organization."""
	domains = [{"name": tenant["domain"], "isDefault": True, "isVerified": True}]
	return {
		"value": [{"id": "test-tenant", "displayName": "Test Mode Tenant", "verifiedDomains": domains}]
	}


def subscribed_skus(tenant, params, body):
	"""GET /subscribedSkus with consumed units counted from users."""
	value = []
	for sku in tenant["skus"].values():
		used = sum(sku["skuId"] in u["licenses"] for u in tenant["users"].values())
		value.append(
			{
				"skuId": sku["skuId"],
				"skuPartNumber": sku["skuPartNumber"],
				"capabilityStatus": "Enabled",
				"consumedUnits": used,
				"prepaidUnits": {"enabled": sku["units"]},
			}
		)
	return {"value": value}


def list_groups(tenant, params, body):
	"""GET /groups, honouring a mailNickname eq filter."""
	nick = re.findall(r"mailNickname eq '((?:[^']|'')*)'", params.get("$filter") or "")
	groups = tenant["groups"].values()
	if nick:
		groups = [g for g in groups if g.get("mailNickname") == nick[0].replace("''", "'")]
	return {"value": [group_reply(g) for g in groups]}


def list_users(tenant, params, body):
	"""GET /users, honouring a userPrincipalName/mail eq filter."""
	wanted = {
		v.replace("''", "'").lower()
		for v in re.findall(r"'((?:[^']|'')*)'", params.get("$filter") or "")
	}
	wanted = {w.removeprefix("smtp:") for w in wanted}
	users = [
		public(u)
		for u in tenant["users"].values()
		if not wanted or u["userPrincipalName"] in wanted or (u.get("mail") or "") in wanted
	]
	return {"value": users}


def create_user(tenant, params, body):
	"""POST /users."""
	if not (body.get("passwordProfile") or {}).get("password"):
		fail(400, "Request_BadRequest", "A password is required.")
	return public(new_user(tenant, body))


def get_user(tenant, key, params, body):
	"""GET /users/{id}."""
	return public(find_user(tenant, key))


def update_user(tenant, key, params, body):
	"""PATCH /users/{id}."""
	find_user(tenant, key).update({k: v for k, v in body.items() if k != "passwordProfile"})
	return {}


def delete_user(tenant, key, params, body):
	"""DELETE /users/{id}."""
	user = find_user(tenant, key)
	for group in tenant["groups"].values():
		group["members"] = [m for m in group["members"] if m != user["id"]]
	del tenant["users"][user["id"]]
	return {}


def set_manager(tenant, key, params, body):
	"""PUT /users/{id}/manager/$ref."""
	manager = find_user(tenant, body.get("@odata.id", "").rsplit("/", 1)[-1])
	find_user(tenant, key)["manager"] = manager["id"]
	return {}


def assign_license(tenant, key, params, body):
	"""POST /users/{id}/assignLicense with seat and usage-location checks."""
	user = find_user(tenant, key)
	adds = [a["skuId"] for a in body.get("addLicenses") or []]
	if adds and not user.get("usageLocation"):
		fail(
			400,
			"Request_BadRequest",
			"License assignment cannot be done for user with invalid usage location.",
		)
	for sku_id in body.get("removeLicenses") or []:
		if sku_id not in user["licenses"]:
			fail(400, "Request_BadRequest", "User does not have a corresponding license.")
	for sku_id in adds:
		sku = tenant["skus"].get(sku_id) or fail(
			400, "Request_BadRequest", f"License {sku_id} does not correspond to a valid company License."
		)
		used = sum(sku_id in u["licenses"] for u in tenant["users"].values())
		if sku_id not in user["licenses"] and used >= sku["units"]:
			fail(
				400,
				"Request_BadRequest",
				f"Subscription {sku['skuPartNumber']} does not have any available licenses.",
			)
	user["licenses"] = [s for s in user["licenses"] if s not in (body.get("removeLicenses") or [])]
	user["licenses"] += [s for s in adds if s not in user["licenses"]]
	return public(user)


def license_details(tenant, key, params, body):
	"""GET /users/{id}/licenseDetails."""
	skus = tenant["skus"]
	licenses = find_user(tenant, key)["licenses"]
	return {"value": [{"skuId": s, "skuPartNumber": skus[s]["skuPartNumber"]} for s in licenses]}


def member_of(tenant, key, params, body):
	"""GET /users/{id}/memberOf/microsoft.graph.group."""
	user_id = find_user(tenant, key)["id"]
	return {"value": [group_reply(g) for g in tenant["groups"].values() if user_id in g["members"]]}


def revoke_sessions(tenant, key, params, body):
	"""POST /users/{id}/revokeSignInSessions."""
	find_user(tenant, key)["sessionsRevoked"] += 1
	return {"value": True}


def send_mail(tenant, key, params, body):
	"""POST /users/{id}/sendMail (kept for inspection; password masked)."""
	find_user(tenant, key)
	message = body.get("message") or {}
	recipients = [r["emailAddress"]["address"] for r in message.get("toRecipients") or []]
	tenant["mail"].append({"from": key, "to": recipients, "subject": message.get("subject")})
	return {}


def add_member(tenant, group_id, params, body):
	"""POST /groups/{id}/members/$ref."""
	group = tenant["groups"].get(group_id) or fail(
		404, "Request_ResourceNotFound", "Group not found."
	)
	if "DynamicMembership" in group["groupTypes"] or (
		group["mailEnabled"] and "Unified" not in group["groupTypes"]
	):
		fail(
			400,
			"Request_BadRequest",
			"Cannot Update a mail-enabled security groups and or distribution list.",
		)
	user_id = find_user(tenant, body.get("@odata.id", "").rsplit("/", 1)[-1])["id"]
	if user_id in group["members"]:
		fail(
			400,
			"Request_BadRequest",
			"One or more added object references already exist for the following modified properties: 'members'.",
		)
	group["members"].append(user_id)
	return {}


def remove_member(tenant, group_id, user_key, params, body):
	"""DELETE /groups/{id}/members/{user}/$ref."""
	group = tenant["groups"].get(group_id) or fail(
		404, "Request_ResourceNotFound", "Group not found."
	)
	user_id = find_user(tenant, user_key)["id"]
	if user_id not in group["members"]:
		fail(
			404,
			"Request_ResourceNotFound",
			f"Resource '{user_id}' does not exist or one of its queried reference-property objects are not present.",
		)
	group["members"].remove(user_id)
	return {}


ID = r"([^/]+)"
ROUTES = [
	(r"/organization", "GET", organization),
	(r"/subscribedSkus", "GET", subscribed_skus),
	(r"/groups", "GET", list_groups),
	(r"/groups", "POST", create_group),
	(rf"/groups/{ID}", "GET", get_group),
	(rf"/groups/{ID}", "DELETE", delete_group),
	(rf"/groups/{ID}/members/microsoft\.graph\.user", "GET", list_members),
	(rf"/groups/{ID}/owners/microsoft\.graph\.user", "GET", list_owners),
	(rf"/groups/{ID}/owners/\$ref", "POST", add_owner),
	(r"/users", "GET", list_users),
	(r"/users", "POST", create_user),
	(rf"/users/{ID}", "GET", get_user),
	(rf"/users/{ID}", "PATCH", update_user),
	(rf"/users/{ID}", "DELETE", delete_user),
	(rf"/users/{ID}/manager/\$ref", "PUT", set_manager),
	(rf"/users/{ID}/assignLicense", "POST", assign_license),
	(rf"/users/{ID}/licenseDetails", "GET", license_details),
	(rf"/users/{ID}/memberOf/microsoft\.graph\.group", "GET", member_of),
	(rf"/users/{ID}/revokeSignInSessions", "POST", revoke_sessions),
	(rf"/users/{ID}/sendMail", "POST", send_mail),
	(rf"/groups/{ID}/members/\$ref", "POST", add_member),
	(rf"/groups/{ID}/members/{ID}/\$ref", "DELETE", remove_member),
]

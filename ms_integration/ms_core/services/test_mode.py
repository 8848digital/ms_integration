# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""
Fake Microsoft 365 tenant for Test Mode.

Answers the Graph calls this app makes with Graph-shaped replies and errors,
keeping users, licenses, groups and sent mails so a demo
behaves like a real tenant. Nothing leaves the server.

The tenant is a JSON file in the site's private folder, so it survives cache
clears and restarts; delete the file (or call reset) to start over.
"""

import json
import os
import re
import uuid

import frappe

FILE_NAME = "microsoft365_test_tenant.json"
SKUS = [
	("6fd2c87f-b296-42f0-b197-1e91e994b900", "ENTERPRISEPACK", 25),
	("f245ecc8-75af-4f8e-b61f-27d8114de5f3", "O365_BUSINESS_PREMIUM", 10),
	("4b9405b0-7788-4568-add1-99614e613b69", "EXCHANGESTANDARD", 2),
]
GROUPS = [
	("All Staff", ["Unified"], True, False),
	("IT Team", [], False, True),
	("Egypt Office", [], False, True),
	("Sales Team", [], True, False),
	("Everyone (Dynamic)", ["Unified", "DynamicMembership"], True, False),
]


def handle(method: str, path: str, params: dict, body: dict | None) -> dict:
	"""
	Answer one Graph call from the fake tenant.

	Parameters:
	        method (str, required): HTTP method.
	        path (str, required): Graph path (or full URL).
	        params (dict, required): Query parameters.
	        body (dict, optional): JSON body.

	Returns:
	        dict: Graph-shaped reply.
	"""
	path = re.sub(r"^https://[^/]+/v1\.0", "", path).split("?")[0]
	tenant = load()
	from ms_integration.ms_core.services.test_mode_routes import ROUTES

	for pattern, verb, handler in ROUTES:
		found = re.fullmatch(pattern, path)
		if found and verb == method:
			reply = handler(tenant, *found.groups(), params=params, body=body or {})
			save(tenant)
			return reply
	fail(400, "BadRequest", f"Test Mode does not support {method} {path}")


def tenant_path() -> str:
	"""
	Path of the fake tenant file.

	Returns:
	        str: File path in the site's private folder.
	"""
	return frappe.get_site_path("private", FILE_NAME)


def load() -> dict:
	"""
	The fake tenant, created on first use.

	Returns:
	        dict: Tenant state.
	"""
	path = tenant_path()
	if not os.path.exists(path):
		return seed()
	with open(path) as file:
		return json.load(file)


def save(tenant: dict) -> None:
	"""
	Persist the fake tenant.

	Parameters:
	        tenant (dict, required): Tenant state.

	Returns:
	        None
	"""
	os.makedirs(os.path.dirname(tenant_path()), exist_ok=True)
	with open(tenant_path(), "w") as file:
		json.dump(tenant, file)


def reset() -> None:
	"""
	Forget the fake tenant (next call starts from the seed).

	Returns:
	        None
	"""
	if os.path.exists(tenant_path()):
		os.remove(tenant_path())


def seed() -> dict:
	"""
	A tenant with three license types, five groups and one admin user.

	Returns:
	        dict: Tenant state.
	"""
	domain = frappe.db.get_single_value("Outlook Settings", "email_domain") or "example.com"
	skus = {sku: {"skuId": sku, "skuPartNumber": part, "units": units} for sku, part, units in SKUS}
	groups = {}
	for name, types, mail, security in GROUPS:
		group_id = str(uuid.uuid5(uuid.NAMESPACE_DNS, name))
		nick = re.sub(r"[^a-z]", "", name.lower())
		groups[group_id] = {
			"id": group_id,
			"displayName": name,
			"groupTypes": types,
			"mailEnabled": mail,
			"securityEnabled": security,
			"mail": f"{nick}@{domain}" if mail else None,
			"mailNickname": nick,
			"description": f"Test Mode group {name}",
			"members": [],
		}
	tenant = {"domain": domain, "skus": skus, "groups": groups, "users": {}, "mail": []}
	new_user(tenant, {"userPrincipalName": f"it.admin@{domain}", "displayName": "IT Admin"})
	return tenant


def fail(status: int, code: str, message: str):
	"""
	Raise a Graph-style error.

	Parameters:
	        status (int, required): HTTP status.
	        code (str, required): Graph error code.
	        message (str, required): Graph error message.
	"""
	from ms_integration.ms_core.services.graph import GraphError

	raise GraphError(status, code, message)


def find_user(tenant: dict, key: str) -> dict:
	"""
	A user by object id or sign-in name; 404 like Graph when missing.

	Parameters:
	        tenant (dict, required): Tenant state.
	        key (str, required): Object id or userPrincipalName.

	Returns:
	        dict: The user.
	"""
	key = (key or "").lower()
	for user in tenant["users"].values():
		if key in (user["id"], user["userPrincipalName"].lower()):
			return user
	fail(404, "Request_ResourceNotFound", f"Resource '{key}' does not exist.")


def new_user(tenant: dict, body: dict) -> dict:
	"""
	Add a user, rejecting a taken sign-in name like Graph does.

	Parameters:
	        tenant (dict, required): Tenant state.
	        body (dict, required): User properties.

	Returns:
	        dict: The new user.
	"""
	upn = (body.get("userPrincipalName") or "").lower()
	if any(u["userPrincipalName"].lower() == upn for u in tenant["users"].values()):
		fail(
			400,
			"Request_BadRequest",
			"Another object with the same value for property userPrincipalName already exists.",
		)
	if not upn.endswith("@" + tenant["domain"].lower()):
		fail(
			400, "Request_BadRequest", "The domain portion of the userPrincipalName property is invalid."
		)
	user = {k: v for k, v in body.items() if k != "passwordProfile"}
	user.update(id=str(uuid.uuid4()), mail=upn, accountEnabled=body.get("accountEnabled", True))
	user.update(userPrincipalName=upn, licenses=[], manager=None, sessionsRevoked=0)
	tenant["users"][user["id"]] = user
	return user


def public(user: dict) -> dict:
	"""
	User as Graph returns it (without fake-only keys).

	Parameters:
	        user (dict, required): Stored user.

	Returns:
	        dict: Graph user.
	"""
	return {k: v for k, v in user.items() if k not in ("licenses", "manager", "sessionsRevoked")}


def group_reply(group: dict) -> dict:
	"""
	Group as Graph returns it.

	Parameters:
	        group (dict, required): Stored group.

	Returns:
	        dict: Graph group.
	"""
	return {
		"@odata.type": "#microsoft.graph.group",
		**{k: v for k, v in group.items() if k not in ("members", "owners")},
	}

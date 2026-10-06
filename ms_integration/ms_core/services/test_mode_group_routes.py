# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Test Mode fake tenant: routes for groups created from documents."""

import uuid

from ms_integration.ms_core.services.test_mode import fail, find_user, group_reply


def create_group(tenant, params, body):
	"""POST /groups (owners bound at creation)."""
	nick = body.get("mailNickname")
	if any(g.get("mailNickname") == nick for g in tenant["groups"].values()):
		fail(
			400,
			"Request_BadRequest",
			"Another object with the same value for property mailNickname already exists.",
		)
	owners = [
		find_user(tenant, ref.rsplit("/", 1)[-1])["id"] for ref in body.get("owners@odata.bind") or []
	]
	group = {k: v for k, v in body.items() if k != "owners@odata.bind"}
	group.update(id=str(uuid.uuid4()), members=[], owners=owners)
	group["mail"] = f"{nick}@{tenant['domain']}" if body.get("mailEnabled") else None
	tenant["groups"][group["id"]] = group
	return group_reply(group)


def find_group(tenant, group_id):
	"""A group by id; 404 like Graph when missing."""
	return tenant["groups"].get(group_id) or fail(
		404, "Request_ResourceNotFound", f"Resource '{group_id}' does not exist."
	)


def get_group(tenant, group_id, params, body):
	"""GET /groups/{id}."""
	return group_reply(find_group(tenant, group_id))


def delete_group(tenant, group_id, params, body):
	"""DELETE /groups/{id}."""
	find_group(tenant, group_id)
	del tenant["groups"][group_id]
	return {}


def users_in(tenant, ids):
	"""Graph users for object ids."""
	users = tenant["users"]
	return {
		"value": [
			{k: users[i][k] for k in ("id", "userPrincipalName", "mail")} for i in ids if i in users
		]
	}


def list_members(tenant, group_id, params, body):
	"""GET /groups/{id}/members/microsoft.graph.user."""
	return users_in(tenant, find_group(tenant, group_id)["members"])


def list_owners(tenant, group_id, params, body):
	"""GET /groups/{id}/owners/microsoft.graph.user."""
	return users_in(tenant, find_group(tenant, group_id).get("owners") or [])


def add_owner(tenant, group_id, params, body):
	"""POST /groups/{id}/owners/$ref."""
	group = find_group(tenant, group_id)
	user_id = find_user(tenant, body.get("@odata.id", "").rsplit("/", 1)[-1])["id"]
	if user_id in group.setdefault("owners", []):
		fail(
			400,
			"Request_BadRequest",
			"One or more added object references already exist for the following modified properties: 'owners'.",
		)
	group["owners"].append(user_id)
	return {}

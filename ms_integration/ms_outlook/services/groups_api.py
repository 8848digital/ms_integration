# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Microsoft Graph operations for groups made from documents."""

import re

from ms_integration.ms_core.services.graph import GraphClient, GraphError, base_url, quote
from ms_integration.ms_outlook.services.directory_api import GROUP_FIELDS

MEMBER_FIELDS = "id,userPrincipalName,mail"


def nickname_base(name: str) -> str:
	"""
	Mail nickname from a group name (letters, digits, dots, dashes; max 50).

	Parameters:
	        name (str, required): Group display name.

	Returns:
	        str: Nickname base.
	"""
	nick = re.sub(r"[^a-z0-9.]+", "-", (name or "").lower())
	nick = re.sub(r"[.-]{2,}", "-", nick).strip(".-")
	return (nick or "group")[:50]


def free_nickname(client: GraphClient, name: str) -> str:
	"""
	A mail nickname no other group uses (adds -2, -3, … when needed).

	Parameters:
	        client (GraphClient, required): Graph client.
	        name (str, required): Group display name.

	Returns:
	        str: Free nickname.
	"""
	base = nickname_base(name)
	for attempt in range(1, 50):
		nick = base if attempt == 1 else f"{base[:46]}-{attempt}"
		found = client.get(
			"/groups", params={"$filter": f"mailNickname eq '{quote(nick)}'", "$select": "id"}
		)
		if not found.get("value"):
			return nick
	raise GraphError(400, "nickname_taken", f"No free mail nickname for {name}")


def create_group(
	client: GraphClient, name: str, kind: str, description: str, visibility: str, owners=()
) -> dict:
	"""
	Create a Microsoft 365 or security group.

	Parameters:
	        client (GraphClient, required): Graph client.
	        name (str, required): Display name.
	        kind (str, required): "Microsoft 365" or "Security".
	        description (str, required): Description.
	        visibility (str, required): Private / Public (Microsoft 365 groups).
	        owners (Iterable[str], optional): Owner object ids set at creation.

	Returns:
	        dict: Created Graph group.
	"""
	unified = kind == "Microsoft 365"
	body = {
		"displayName": name,
		"description": description or name,
		"mailNickname": free_nickname(client, name),
		"mailEnabled": unified,
		"securityEnabled": not unified,
		"groupTypes": ["Unified"] if unified else [],
	}
	if unified:
		body["visibility"] = visibility or "Private"
	if owners:
		body["owners@odata.bind"] = [f"{base_url()}/users/{owner}" for owner in owners]
	return client.request("POST", "/groups", json=body)


def get_group(client: GraphClient, group_id: str) -> dict:
	"""
	One group.

	Parameters:
	        client (GraphClient, required): Graph client.
	        group_id (str, required): Group object id.

	Returns:
	        dict: Graph group.
	"""
	return client.get(f"/groups/{group_id}", params={"$select": GROUP_FIELDS})


def delete_group(client: GraphClient, group_id: str) -> None:
	"""
	Delete a group (kept 30 days in Microsoft's recycle bin; 404 = already gone).

	Parameters:
	        client (GraphClient, required): Graph client.
	        group_id (str, required): Group object id.

	Returns:
	        None
	"""
	try:
		client.request("DELETE", f"/groups/{group_id}")
	except GraphError as e:
		if e.status != 404:
			raise


def group_members(client: GraphClient, group_id: str) -> list[dict]:
	"""
	Users who are direct members.

	Parameters:
	        client (GraphClient, required): Graph client.
	        group_id (str, required): Group object id.

	Returns:
	        list[dict]: Graph users.
	"""
	path = f"/groups/{group_id}/members/microsoft.graph.user"
	return client.get_all(path, params={"$select": MEMBER_FIELDS})


def group_owners(client: GraphClient, group_id: str) -> list[dict]:
	"""
	Users who are owners.

	Parameters:
	        client (GraphClient, required): Graph client.
	        group_id (str, required): Group object id.

	Returns:
	        list[dict]: Graph users.
	"""
	path = f"/groups/{group_id}/owners/microsoft.graph.user"
	return client.get_all(path, params={"$select": MEMBER_FIELDS})


def add_owner(client: GraphClient, group_id: str, user_id: str) -> None:
	"""
	Make a user an owner (already an owner is fine).

	Parameters:
	        client (GraphClient, required): Graph client.
	        group_id (str, required): Group object id.
	        user_id (str, required): User object id.

	Returns:
	        None
	"""
	reference = {"@odata.id": f"{base_url()}/users/{user_id}"}
	try:
		client.request("POST", f"/groups/{group_id}/owners/$ref", json=reference)
	except GraphError as e:
		if not (e.status == 400 and "already exist" in e.message):
			raise

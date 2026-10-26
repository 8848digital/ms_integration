# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Microsoft Graph user, license, group and mail operations used by the app."""

from ms_integration.ms_core.services.graph import GraphClient, GraphError, base_url, quote

USER_FIELDS = "id,userPrincipalName,mail,displayName,accountEnabled"
GROUP_FIELDS = "id,displayName,mail,mailEnabled,securityEnabled,groupTypes,description"
ALREADY_MEMBER = "already exist"


def find_user_by_address(client: GraphClient, address: str) -> dict | None:
	"""
	The user whose sign-in name, primary email or alias is `address`.

	Parameters:
	        client (GraphClient, required): Graph client.
	        address (str, required): Email address.

	Returns:
	        dict | None: Graph user, or None when the address is free.
	"""
	value = quote(address.lower())
	condition = (
		f"userPrincipalName eq '{value}' or mail eq '{value}' "
		f"or proxyAddresses/any(p:p eq 'smtp:{value}')"
	)
	users = client.get(
		"/users",
		params={"$filter": condition, "$select": USER_FIELDS, "$count": "true"},
		headers={"ConsistencyLevel": "eventual"},
	).get("value")
	return users[0] if users else None


def all_users(client: GraphClient) -> list[dict]:
	"""
	Every user in the tenant (id, sign-in name, mail, enabled).

	Parameters:
	        client (GraphClient, required): Graph client.

	Returns:
	        list[dict]: Graph users.
	"""
	return client.get_all("/users", params={"$select": USER_FIELDS, "$top": "999"})


def create_user(client: GraphClient, payload: dict) -> dict:
	"""
	Create a user.

	Parameters:
	        client (GraphClient, required): Graph client.
	        payload (dict, required): Graph user body including passwordProfile.

	Returns:
	        dict: Created user.
	"""
	return client.request("POST", "/users", json=payload)


def update_user(client: GraphClient, user_id: str, payload: dict) -> None:
	"""
	Update user properties.

	Parameters:
	        client (GraphClient, required): Graph client.
	        user_id (str, required): Object id or sign-in name.
	        payload (dict, required): Properties to change.

	Returns:
	        None
	"""
	if payload:
		client.request("PATCH", f"/users/{user_id}", json=payload)


def set_manager(client: GraphClient, user_id: str, manager_id: str) -> None:
	"""
	Set the user's manager.

	Parameters:
	        client (GraphClient, required): Graph client.
	        user_id (str, required): User object id.
	        manager_id (str, required): Manager object id.

	Returns:
	        None
	"""
	reference = {"@odata.id": f"{base_url()}/users/{manager_id}"}
	client.request("PUT", f"/users/{user_id}/manager/$ref", json=reference)


def assign_licenses(client: GraphClient, user_id: str, add=(), remove=()) -> None:
	"""
	Add and remove licenses in one atomic call.

	Parameters:
	        client (GraphClient, required): Graph client.
	        user_id (str, required): User object id.
	        add (Iterable[str], optional): SKU ids to add.
	        remove (Iterable[str], optional): SKU ids to remove.

	Returns:
	        None
	"""
	if not add and not remove:
		return
	body = {
		"addLicenses": [{"skuId": sku, "disabledPlans": []} for sku in add],
		"removeLicenses": list(remove),
	}
	client.request("POST", f"/users/{user_id}/assignLicense", json=body)


def user_license_skus(client: GraphClient, user_id: str) -> list[str]:
	"""
	SKU ids assigned to the user.

	Parameters:
	        client (GraphClient, required): Graph client.
	        user_id (str, required): User object id.

	Returns:
	        list[str]: SKU ids.
	"""
	return [row["skuId"] for row in client.get_all(f"/users/{user_id}/licenseDetails")]


def user_groups(client: GraphClient, user_id: str) -> list[dict]:
	"""
	Groups the user is a direct member of.

	Parameters:
	        client (GraphClient, required): Graph client.
	        user_id (str, required): User object id.

	Returns:
	        list[dict]: Graph groups.
	"""
	path = f"/users/{user_id}/memberOf/microsoft.graph.group"
	return client.get_all(path, params={"$select": GROUP_FIELDS})


def add_to_group(client: GraphClient, group_id: str, user_id: str) -> bool:
	"""
	Add the user to a group.

	Parameters:
	        client (GraphClient, required): Graph client.
	        group_id (str, required): Group object id.
	        user_id (str, required): User object id.

	Returns:
	        bool: False when the user was already a member.
	"""
	reference = {"@odata.id": f"{base_url()}/directoryObjects/{user_id}"}
	try:
		client.request("POST", f"/groups/{group_id}/members/$ref", json=reference)
	except GraphError as e:
		if e.status == 400 and ALREADY_MEMBER in e.message:
			return False
		raise
	return True


def remove_from_group(client: GraphClient, group_id: str, user_id: str) -> bool:
	"""
	Remove the user from a group.

	Parameters:
	        client (GraphClient, required): Graph client.
	        group_id (str, required): Group object id.
	        user_id (str, required): User object id.

	Returns:
	        bool: False when the user wasn't a member.
	"""
	try:
		client.request("DELETE", f"/groups/{group_id}/members/{user_id}/$ref")
	except GraphError as e:
		if e.status == 404:
			return False
		raise
	return True

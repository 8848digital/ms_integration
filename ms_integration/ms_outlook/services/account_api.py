# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Microsoft Graph sign-in, password, deletion and mail operations."""

from ms_integration.ms_core.services.graph import GraphClient


def set_sign_in(client: GraphClient, user_id: str, enabled: bool) -> None:
	"""
	Allow or block sign-in.

	Parameters:
	        client (GraphClient, required): Graph client.
	        user_id (str, required): User object id.
	        enabled (bool, required): Allow sign-in.

	Returns:
	        None
	"""
	client.request("PATCH", f"/users/{user_id}", json={"accountEnabled": enabled})


def revoke_sessions(client: GraphClient, user_id: str) -> None:
	"""
	Sign the user out everywhere (refresh tokens invalidated).

	Parameters:
	        client (GraphClient, required): Graph client.
	        user_id (str, required): User object id.

	Returns:
	        None
	"""
	client.request("POST", f"/users/{user_id}/revokeSignInSessions")


def set_password(client: GraphClient, user_id: str, password: str) -> None:
	"""
	Set a temporary password that must be changed at next sign-in.

	Parameters:
	        client (GraphClient, required): Graph client.
	        user_id (str, required): User object id.
	        password (str, required): New password.

	Returns:
	        None
	"""
	profile = {"forceChangePasswordNextSignIn": True, "password": password}
	client.request("PATCH", f"/users/{user_id}", json={"passwordProfile": profile})


def delete_user(client: GraphClient, user_id: str) -> None:
	"""
	Delete the user (kept 30 days in Microsoft's recycle bin).

	Parameters:
	        client (GraphClient, required): Graph client.
	        user_id (str, required): User object id.

	Returns:
	        None
	"""
	client.request("DELETE", f"/users/{user_id}")


def send_mail(client: GraphClient, sender: str, to: str, subject: str, html: str) -> None:
	"""
	Send an HTML email from a mailbox in the tenant (not saved in Sent Items).

	Parameters:
	        client (GraphClient, required): Graph client.
	        sender (str, required): Sender mailbox.
	        to (str, required): Recipient.
	        subject (str, required): Subject.
	        html (str, required): HTML body.

	Returns:
	        None
	"""
	message = {
		"subject": subject,
		"body": {"contentType": "HTML", "content": html},
		"toRecipients": [{"emailAddress": {"address": to}}],
	}
	client.request(
		"POST", f"/users/{sender}/sendMail", json={"message": message, "saveToSentItems": False}
	)

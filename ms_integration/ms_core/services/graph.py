# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""
Microsoft Graph client (app-only, client credentials).

One cached token per tenant/app, retries on throttling (429/503/504 with
Retry-After), follows @odata.nextLink paging. In Test Mode every call goes to
the built-in fake tenant instead.
"""

import hashlib
import time

import requests

import frappe
from frappe import _

from ms_integration.ms_core.services import test_mode
from ms_integration.ms_core.services.settings import get_connection

GRAPH_URL = "https://graph.microsoft.com/v1.0"
LOGIN_URL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
SCOPE = "https://graph.microsoft.com/.default"
RETRY_STATUSES = (429, 503, 504)
MAX_ATTEMPTS = 4
MAX_WAIT = 10


class GraphError(Exception):
	"""A Microsoft Graph call that failed, with Microsoft's status, code and message."""

	def __init__(self, status: int, code: str, message: str):
		"""
		Parameters:
		        status (int, required): HTTP status.
		        code (str, required): Graph error code, e.g. Request_ResourceNotFound.
		        message (str, required): Graph error message.
		"""
		super().__init__(f"{status} {code}: {message}")
		self.status, self.code, self.message = status, code, message


class GraphClient:
	"""Calls Microsoft Graph with the credentials in Microsoft 365 Settings."""

	def __init__(self):
		"""Read the settings once per client."""
		self.settings = get_connection()
		self.timeout = self.settings.timeout or 30
		self.test_mode = bool(self.settings.test_mode)

	def get(self, path: str, params: dict | None = None, **kw) -> dict:
		"""
		GET a Graph resource.

		Parameters:
		        path (str, required): Path after /v1.0, e.g. `/users/{id}`.
		        params (dict, optional): Query parameters.

		Returns:
		        dict: Response JSON.
		"""
		return self.request("GET", path, params=params, **kw)

	def get_all(self, path: str, params: dict | None = None, **kw) -> list[dict]:
		"""
		GET every page of a collection (follows @odata.nextLink).

		Parameters:
		        path (str, required): Collection path.
		        params (dict, optional): Query parameters for the first page.

		Returns:
		        list[dict]: All items.
		"""
		items, page = [], self.request("GET", path, params=params, **kw)
		while True:
			items.extend(page.get("value") or [])
			next_link = page.get("@odata.nextLink")
			if not next_link:
				return items
			page = self.request("GET", next_link, **kw)

	def request(self, method: str, path: str, params=None, json=None, headers=None) -> dict:
		"""
		Call Graph and return the JSON body ({} for empty replies).

		Parameters:
		        method (str, required): HTTP method.
		        path (str, required): Path after /v1.0 or a full nextLink URL.
		        params (dict, optional): Query parameters.
		        json (dict, optional): JSON body.
		        headers (dict, optional): Extra headers (e.g. ConsistencyLevel).

		Returns:
		        dict: Response JSON.

		Raises:
		        GraphError: On any non-2xx reply after retries.
		"""
		if self.test_mode:
			return test_mode.handle(method, path, params or {}, json)
		url = path if path.startswith("https://") else f"{base_url()}{path}"
		for attempt in range(1, MAX_ATTEMPTS + 1):
			response = requests.request(
				method,
				url,
				params=params,
				json=json,
				headers={"Authorization": f"Bearer {self.token()}", **(headers or {})},
				timeout=self.timeout,
			)
			if response.status_code not in RETRY_STATUSES or attempt == MAX_ATTEMPTS:
				break
			time.sleep(min(int(response.headers.get("Retry-After") or attempt), MAX_WAIT))
		if response.status_code >= 400:
			raise error_from(response)
		return response.json() if response.content else {}

	def token(self) -> str:
		"""
		Access token, cached until five minutes before it expires.

		Returns:
		        str: Bearer token.
		"""
		key = token_cache_key(self.settings)
		token = frappe.cache.get_value(key)
		if token:
			return token
		token, expires_in = self.__fetch_token()
		frappe.cache.set_value(key, token, expires_in_sec=max(expires_in - 300, 60))
		return token

	def __fetch_token(self) -> tuple[str, int]:
		"""
		Client-credentials token from Microsoft Entra ID.

		Returns:
		        tuple[str, int]: Token and lifetime in seconds.
		"""
		settings = self.settings
		response = requests.post(
			LOGIN_URL.format(tenant=settings.tenant_id),
			data={
				"grant_type": "client_credentials",
				"client_id": settings.client_id,
				"client_secret": settings.get_password("client_secret", raise_exception=False),
				"scope": SCOPE,
			},
			timeout=self.timeout,
		)
		body = response.json() if response.content else {}
		if response.status_code != 200 or not body.get("access_token"):
			description = (body.get("error_description") or response.text or "").split("\r\n")[0]
			raise GraphError(response.status_code, body.get("error") or "token_error", description)
		return body["access_token"], int(body.get("expires_in") or 3600)


def base_url() -> str:
	"""
	Graph base URL; `microsoft_graph_base_url` in site_config overrides it (local fakes).

	Returns:
	        str: Base URL without trailing slash.
	"""
	return (frappe.conf.get("microsoft_graph_base_url") or GRAPH_URL).rstrip("/")


def token_cache_key(settings) -> str:
	"""
	Cache key bound to tenant + client, so changed credentials get a new token.

	Parameters:
	        settings (Document, required): Microsoft 365 Settings.

	Returns:
	        str: Cache key.
	"""
	ident = f"{settings.tenant_id}:{settings.client_id}".encode()
	return f"microsoft365_token:{hashlib.sha256(ident).hexdigest()[:16]}"


def error_from(response) -> GraphError:
	"""
	GraphError from a failed response.

	Parameters:
	        response (requests.Response, required): Failed response.

	Returns:
	        GraphError: Parsed error.
	"""
	try:
		error = response.json().get("error") or {}
	except ValueError:
		error = {}
	return GraphError(
		response.status_code, error.get("code") or "error", error.get("message") or response.text[:500]
	)


def friendly(error: Exception) -> str:
	"""
	Short message for users and logs.

	Parameters:
	        error (Exception, required): Any error.

	Returns:
	        str: Message.
	"""
	if isinstance(error, GraphError):
		return _("Microsoft 365 said: {0} ({1})").format(error.message, error.code)
	return str(error)


def quote(value: str) -> str:
	"""
	Escape a value for an OData $filter string literal.

	Parameters:
	        value (str, required): Raw value.

	Returns:
	        str: Value with single quotes doubled.
	"""
	return (value or "").replace("'", "''")

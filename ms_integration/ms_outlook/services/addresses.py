# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""
New email addresses from the configured pattern.

`{last}` starts with N letters of the last name and grows one letter at a time
until the address is free in ERPNext and in Microsoft 365, e.g.
john.s → john.sm → john.smi. When the whole last name is taken, it fails so
IT can type an address by hand.
"""

import re
import unicodedata

import frappe
from frappe import _

from ms_integration.ms_core.services.settings import lines
from ms_integration.ms_outlook.services.directory_api import find_user_by_address

LOCAL_PART = re.compile(r"^[a-z0-9]+([._-][a-z0-9]+)*$")


class NoFreeAddress(frappe.ValidationError):
	"""Every address the pattern can build is already taken."""


def clean(text: str | None) -> str:
	"""
	Lower-case ASCII letters and digits only (accents removed, spaces dropped).

	Parameters:
	        text (str, optional): Name part.

	Returns:
	        str: Cleaned text.
	"""
	ascii_text = unicodedata.normalize("NFKD", text or "").encode("ascii", "ignore").decode()
	return re.sub(r"[^a-z0-9]", "", ascii_text.lower())


def last_name_base(last_name: str | None, prefixes: list[str]) -> str:
	"""
	Last name without leading prefixes written as separate words or hyphenated.

	"Al Harbi" / "Al-Harbi" → "harbi"; "Alharbi" and "Ellis" stay unchanged.

	Parameters:
	        last_name (str, optional): Employee last name.
	        prefixes (list[str], required): Prefixes to drop, lower case.

	Returns:
	        str: Cleaned last name.
	"""
	words = [w for w in re.split(r"[\s\-]+", last_name or "") if w]
	while len(words) > 1 and clean(words[0]) in prefixes:
		words.pop(0)
	return clean("".join(words))


def candidates(employee, settings):
	"""
	Addresses to try, in order.

	Parameters:
	        employee (Document, required): Employee.
	        settings (Document, required): Microsoft 365 Settings.

	Yields:
	        str: Email address.
	"""
	pattern = (settings.email_pattern or "{first}.{last}").strip().lower()
	first = clean(employee.first_name)
	last = last_name_base(employee.last_name, [clean(p) for p in lines(settings.drop_prefixes)])
	tokens = {"{first}": first, "{middle}": clean(employee.middle_name), "{first_initial}": first[:1]}
	base = pattern
	for token, value in tokens.items():
		base = base.replace(token, value)
	if "{last}" not in base:
		yield from address(base, settings.email_domain)
		return
	start = settings.min_last_letters or len(last)
	for letters in range(max(min(start, len(last)), 1), len(last) + 1):
		yield from address(base.replace("{last}", last[:letters]), settings.email_domain)


def address(local: str, domain: str):
	"""
	The full address when the local part is valid.

	Parameters:
	        local (str, required): Local part, possibly with stray separators.
	        domain (str, required): Email domain.

	Yields:
	        str: Address (nothing when invalid).
	"""
	local = re.sub(r"([._-])[._-]+", r"\1", local).strip("._-")
	if LOCAL_PART.match(local):
		yield f"{local}@{(domain or '').strip().lower()}"


def taken_in_erpnext(email: str, employee: str, email_field: str | None) -> bool:
	"""
	Whether ERPNext already uses the address for someone else.

	Parameters:
	        email (str, required): Address.
	        employee (str, required): Employee being provisioned.
	        email_field (str, optional): Employee field holding the work email.

	Returns:
	        bool: True when a User, another Employee or another account has it.
	"""
	if frappe.db.exists("User", {"name": email}) or frappe.db.exists("User", {"email": email}):
		return True
	others = {"name": ("!=", employee)}
	if email_field and frappe.db.exists("Employee", {email_field: email, **others}):
		return True
	return bool(
		frappe.db.exists(
			"Microsoft 365 Account", {"user_principal_name": email, "employee": ("!=", employee)}
		)
	)


def is_free(client, email: str, employee: str, settings) -> bool:
	"""
	Free in ERPNext and in Microsoft 365.

	Parameters:
	        client (GraphClient, required): Graph client.
	        email (str, required): Address.
	        employee (str, required): Employee being provisioned.
	        settings (Document, required): Microsoft 365 Settings.

	Returns:
	        bool: True when nobody has it.
	"""
	if taken_in_erpnext(email, employee, settings.email_field):
		return False
	return find_user_by_address(client, email) is None


def choose_address(client, employee, settings, skip=()) -> str:
	"""
	First free address for the employee.

	Parameters:
	        client (GraphClient, required): Graph client.
	        employee (Document, required): Employee.
	        settings (Document, required): Microsoft 365 Settings.
	        skip (Iterable[str], optional): Addresses already found taken.

	Returns:
	        str: Address.

	Raises:
	        NoFreeAddress: When every candidate is taken.
	"""
	for email in candidates(employee, settings):
		if email not in skip and is_free(client, email, employee.name, settings):
			return email
	raise NoFreeAddress(
		_(
			"No free email address could be built for {0}. Type one in the Microsoft 365 Account and click Retry."
		).format(employee.employee_name or employee.name)
	)

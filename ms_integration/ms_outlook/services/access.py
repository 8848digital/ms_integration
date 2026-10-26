# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""
Licenses and groups: what the Provisioning Rules give an employee, and
pushing edits of a Microsoft 365 Account's rows to Microsoft 365.
"""

import frappe
from frappe import _

from ms_integration.ms_core.services.graph import GraphClient, friendly
from ms_integration.ms_core.services.logs import log, summary
from ms_integration.ms_core.services.settings import matches
from ms_integration.ms_outlook.services import directory_api as api
from ms_integration.ms_outlook.services.settings import is_enabled

GROUP = "Microsoft 365 Group"


def rule_access(employee) -> tuple[list[str], list[str]]:
	"""
	Licenses and groups of every enabled rule whose condition matches.

	Parameters:
	        employee (Document, required): Employee.

	Returns:
	        tuple[list[str], list[str]]: License names and group names.
	"""
	licenses, groups = [], []
	for name in frappe.get_all(
		"Microsoft 365 Provisioning Rule", {"enabled": 1}, pluck="name", order_by="name"
	):
		rule = frappe.get_cached_doc("Microsoft 365 Provisioning Rule", name)
		if matches(rule.condition, employee):
			licenses += [row.license for row in rule.licenses if row.license not in licenses]
			groups += [row.group for row in rule.groups if row.group not in groups]
	return licenses, groups


def grant(client: GraphClient, account, licenses: list[str], groups: list[str]) -> list[str]:
	"""
	Add licenses and groups to an account in Microsoft 365 (best effort, logged).

	Parameters:
	        client (GraphClient, required): Graph client.
	        account (Document, required): Microsoft 365 Account with object_id.
	        licenses (list[str], required): License names.
	        groups (list[str], required): Group names.

	Returns:
	        list[str]: Error messages (empty when all succeeded).
	"""
	errors = []
	if licenses:
		try:
			api.assign_licenses(client, account.object_id, add=sku_ids(licenses))
			log("Licenses", account=account.name, details=summary(licenses, label="licenses"))
		except Exception as e:
			errors.append(_("Licenses {0}: {1}").format(", ".join(licenses), friendly(e)))
			log("Licenses", False, account.name, summary(licenses, label="licenses"), e)
	for group in groups:
		try:
			api.add_to_group(client, group, account.object_id)
			log("Groups", account=account.name, details=summary([group_label(group)], label="groups"))
		except Exception as e:
			errors.append(_("Group {0}: {1}").format(group_label(group), friendly(e)))
			log("Groups", False, account.name, summary([group_label(group)], label="groups"), e)
	return errors


def push_changes(account) -> None:
	"""
	On save of an active account, add/remove the licenses and groups that were
	added/removed in its rows. Rows that fail are put back as they were so the
	account always shows the real state; failures are shown and logged.

	Parameters:
	        account (Document, required): Microsoft 365 Account being saved.

	Returns:
	        None
	"""
	before = account.get_doc_before_save()
	if account.flags.from_sync or not before or account.status != "Active" or not is_enabled():
		return
	old_l, new_l = names(before.licenses, "license"), names(account.licenses, "license")
	old_g, new_g = names(before.groups, "group"), names(account.groups, "group")
	if old_l == new_l and old_g == new_g:
		return
	check_groups_manageable(new_g ^ old_g)
	client = GraphClient()
	errors = push_licenses(client, account, new_l - old_l, old_l - new_l)
	errors += push_groups(client, account, new_g - old_g, old_g - new_g)
	if errors:
		frappe.msgprint("<br>".join(errors), title=_("Some changes were not made"), indicator="orange")
	else:
		frappe.msgprint(_("Updated in Microsoft 365."), alert=True, indicator="green")


def push_licenses(client: GraphClient, account, add: set, remove: set) -> list[str]:
	"""
	One atomic license change; on failure the license rows are restored.

	Parameters:
	        client (GraphClient, required): Graph client.
	        account (Document, required): Microsoft 365 Account.
	        add (set, required): License names added.
	        remove (set, required): License names removed.

	Returns:
	        list[str]: Error messages.
	"""
	if not add and not remove:
		return []
	details = summary(sorted(add), sorted(remove), "licenses")
	try:
		api.assign_licenses(client, account.object_id, sku_ids(add), sku_ids(remove))
	except Exception as e:
		kept = [row.license for row in account.licenses if row.license not in add]
		account.set("licenses", [{"license": name} for name in kept + sorted(remove)])
		log("Licenses", False, account.name, details, e)
		return [_("Licenses not changed: {0}").format(friendly(e))]
	log("Licenses", account=account.name, details=details)
	return []


def push_groups(client: GraphClient, account, add: set, remove: set) -> list[str]:
	"""
	Add/remove each group; a failed row is restored.

	Parameters:
	        client (GraphClient, required): Graph client.
	        account (Document, required): Microsoft 365 Account.
	        add (set, required): Group names added.
	        remove (set, required): Group names removed.

	Returns:
	        list[str]: Error and notice messages.
	"""
	messages, done_add, done_remove = [], [], []
	for group in sorted(add):
		try:
			if not api.add_to_group(client, group, account.object_id):
				messages.append(_("Already a member of {0}.").format(group_label(group)))
			done_add.append(group_label(group))
		except Exception as e:
			account.groups = [row for row in account.groups if row.group != group]
			messages.append(_("Could not add {0}: {1}").format(group_label(group), friendly(e)))
	for group in sorted(remove):
		try:
			if not api.remove_from_group(client, group, account.object_id):
				messages.append(_("Was not a member of {0}.").format(group_label(group)))
			done_remove.append(group_label(group))
		except Exception as e:
			account.append("groups", {"group": group})
			messages.append(_("Could not remove {0}: {1}").format(group_label(group), friendly(e)))
	failed = [m for m in messages if "Could not" in m]
	log(
		"Groups",
		not failed,
		account.name,
		summary(done_add, done_remove, "groups"),
		"\n".join(failed) or None,
	)
	return failed


def check_groups_manageable(groups: set) -> None:
	"""
	Stop when a changed group can't have its members changed through Graph.

	Parameters:
	        groups (set, required): Group names being added or removed.

	Returns:
	        None
	"""
	for group in groups:
		if not frappe.db.get_value(GROUP, group, "membership_managed"):
			frappe.throw(
				_(
					"Members of {0} can't be changed from ERPNext (distribution list, mail-enabled security or dynamic group). Change it in Microsoft 365."
				).format(group_label(group))
			)


def names(rows, field: str) -> set[str]:
	"""
	Set of link values in Table MultiSelect rows.

	Parameters:
	        rows (list, required): Child rows.
	        field (str, required): Link field.

	Returns:
	        set[str]: Values.
	"""
	return {row.get(field) for row in rows or [] if row.get(field)}


def sku_ids(licenses) -> list[str]:
	"""
	SKU ids of Microsoft 365 License records.

	Parameters:
	        licenses (Iterable[str], required): License names.

	Returns:
	        list[str]: SKU ids.
	"""
	return [frappe.db.get_value("Microsoft 365 License", name, "sku_id") for name in licenses]


def group_label(group: str) -> str:
	"""
	Display name of a group record.

	Parameters:
	        group (str, required): Group name (object id).

	Returns:
	        str: Display name or id.
	"""
	return frappe.db.get_value(GROUP, group, "display_name") or group

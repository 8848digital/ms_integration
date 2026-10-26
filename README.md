<!--
Copyright (c) 2026 8848 Digital LLP. All rights reserved.
Proprietary and confidential. Unauthorized copying, distribution, or use
of this file, via any medium, is strictly prohibited without prior
written permission from 8848 Digital LLP.
-->

# MS Integration

Microsoft 365 integrations for Frappe/ERPNext in one app, one module per
product, sharing a single Microsoft 365 connection:

- **Outlook** (module *MS Outlook*) — described below.
- **SharePoint** — planned as a separate module in this app.

## Outlook

Employee onboarding and offboarding in Microsoft 365 (Outlook, Teams, OneDrive…)
straight from ERPNext. When an employee becomes active, their Microsoft 365
account is created with the right licenses and groups and the sign-in details
are emailed to them; when they leave, licenses and groups are removed and
sign-in is blocked. IT no longer has to do it by hand in the Microsoft 365
admin center.

Works with any Microsoft 365 tenant and any ERPNext site (HRMS optional).
Everything is configured on screen: no client-specific code.

### Features

- **Account on activation** — when the Employee matches a condition (default
  `status == "Active"`) or with the **Create Account** button.
- **Email address from a pattern** — e.g. `{first}.{last}` with the last name
  growing one letter at a time (`john.s`, `john.sm`, …) until the address is free
  in **both ERPNext and Microsoft 365**; configurable domain and name prefixes to
  drop (e.g. `Al`, `El`). If nothing is free, the account is marked Failed and IT
  types an address and clicks Retry.
- **Secure first sign-in** — a random password per person, changed at first
  sign-in, emailed once to the personal email (through Microsoft 365 or ERPNext
  email). The password is never stored in ERPNext.
- **Provisioning Rules** — conditions on the Employee → licenses and groups
  (e.g. everyone gets E3 + All Staff; `branch == "Egypt"` also gets the Egypt
  Office group). All matching rules apply.
- **Account details kept in sync** — names, job title, department, phone,
  employee ID, company, extension attributes… via a field mapping, plus the
  manager; updated whenever the Employee changes.
- **Licenses and groups from ERPNext** — each employee's **Microsoft 365
  Account** shows their licenses and groups; add or remove a row and save to
  change it in Microsoft 365. Failed changes are put back and explained (e.g. no
  seats left).
- **Offboarding** — when the Employee matches the disable condition (default
  `Inactive` or `Left`): remove licenses, remove from groups, block sign-in and
  sign out of all sessions. Optional deletion after N days.
- **Rehire** — when the employee is active again, sign-in is allowed, the rule
  licenses and groups come back and a new password is emailed.
- **Directory sync** — daily (and on demand): license types with free seats,
  all groups (with paging), each account's licenses and groups, and linking
  employees who already have a mailbox.
- **A group per document (any DocType)** — Group Rules create a Microsoft 365
  group (mail, calendar, Teams-ready) or a security group for each Project,
  Opportunity, Email Group… when a condition matches. Members and owners come
  from the document's fields (Users, Employees, emails — through child tables and
  links, e.g. `users.user`) and from linked documents (e.g. Email Group Members,
  respecting unsubscribe). People removed from the document leave the group; when
  the document closes, members are removed or the group is deleted; the group ID
  and email can be saved on the document for other apps (e.g. SharePoint access).
- **Microsoft 365 Log** — every change and every failure, with Microsoft's
  error message.
- **Test Mode** — a built-in fake tenant to try everything without a Microsoft
  account.

## Modules

- **MS Core** — the shared Microsoft 365 connection (tenant, app credentials,
  Test Mode), the Microsoft Graph client, the Test Mode tenant and the log.
- **MS Outlook** — accounts, onboarding/offboarding, provisioning rules,
  licenses and groups, directory sync, groups per document and the Employee
  customization.

## Setup

See [SETUP.md](SETUP.md) for the Azure app registration, permissions and
settings, and [docs/DEMO.md](docs/DEMO.md) for a demo script.

## License

Proprietary — see [license.txt](license.txt).

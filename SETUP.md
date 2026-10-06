<!--
Copyright (c) 2026 8848 Digital LLP. All rights reserved.
Proprietary and confidential. Unauthorized copying, distribution, or use
of this file, via any medium, is strictly prohibited without prior
written permission from 8848 Digital LLP.
-->

# SETUP.md

## Microsoft 365 (Microsoft Graph)

### Overview
The app calls Microsoft Graph as an **application** (no signed-in user) with a
client secret. Microsoft never calls the site, so no public URL or webhook is
needed.

### Register the app in Microsoft Entra ID (10 minutes, tenant admin)
1. <https://entra.microsoft.com> → **Identity → Applications → App registrations → New registration**.
   Name e.g. `ERPNext Provisioning`, single tenant, no redirect URI → **Register**.
2. Copy **Application (client) ID** and **Directory (tenant) ID**.
3. **Certificates & secrets → New client secret** → copy the **Value** (shown once).
   Note the expiry date — renew it before it expires.
4. **API permissions → Add a permission → Microsoft Graph → Application permissions**:

   | Permission | Why |
   | ---------- | --- |
   | `User.ReadWrite.All` | Create users, update details, block sign-in, sign out |
   | `Group.Read.All` | Read groups for the directory sync (`Group.ReadWrite.All` instead when Group Rules are used: create / delete groups, owners) |
   | `GroupMember.ReadWrite.All` | Add / remove group members |
   | `LicenseAssignment.ReadWrite.All` | Assign / remove licenses (or `User.ReadWrite.All` + `Organization.Read.All` on older tenants) |
   | `Organization.Read.All` | License types and seats, Test Connection |
   | `Mail.Send` | Welcome email from the sender mailbox (only if sent through Microsoft 365) |
   | `User-PasswordProfile.ReadWrite.All` | Reset Password and rehire (new password) |
   | `User.DeleteRestore.All` | Only if "Delete Disabled Accounts After (Days)" is used |

   Then **Grant admin consent**.
5. Recommended: limit `Mail.Send` to the sender mailbox with an Exchange
   **application access policy / RBAC for Applications**, so the app can't send as anyone else.
6. If the tenant blocks password changes for app-only callers, also assign the
   app the **User Administrator** directory role.

### Required Credentials
| Credential | Where it's stored | Required? |
| ---------- | ------------------ | --------- |
| Tenant ID | Microsoft 365 Settings → `tenant_id` | Yes (unless Test Mode) |
| Client ID | Microsoft 365 Settings → `client_id` | Yes (unless Test Mode) |
| Client Secret | Microsoft 365 Settings → `client_secret` (Password, encrypted) | Yes (unless Test Mode) |

### Settings DocType Fields
**Microsoft 365 Settings** (module *MS Core*) — the connection shared by every module

| Field | Type | Mandatory | Notes |
| ----- | ---- | --------- | ----- |
| Enabled | Check | — | Master switch for the Microsoft 365 connection. |
| Test Mode | Check | — | Built-in fake tenant; nothing is sent to Microsoft. Demos/testing only. |
| Timeout (Seconds) | Int | — | Default 30. |
| Tenant ID / Client ID / Client Secret | Data / Data / Password | Yes (live) | From the app registration. |

**Outlook Settings** (module *MS Outlook*) — Outlook tab

| Field | Type | Mandatory | Notes |
| ----- | ---- | --------- | ----- |
| Enabled | Check | — | Turns on the Outlook features (needs the connection enabled too). |
| Sync Licenses, Groups and Accounts Daily | Check | — | Default on. |
| Link Existing Microsoft 365 Accounts | Check | — | During sync, employees whose work email already exists get an account record. |

**Outlook Settings** — New Accounts

| Field | Type | Mandatory | Notes |
| ----- | ---- | --------- | ----- |
| Create Accounts Automatically | Check | — | Off = only the Create Account button. |
| Create When | Code | — | Python on the Employee, default `doc.status == "Active"`. Also re-enables a disabled account. |
| Required Employee Fields | Small Text | — | One per line, default `first_name`, `last_name`. |
| Email Domain | Data | Yes | A verified domain of the tenant. Test Connection warns if it isn't. |
| Address Pattern | Data | — | Default `{first}.{last}`; tokens `{first}`, `{middle}`, `{last}`, `{first_initial}`. |
| Letters of Last Name to Start With | Int | — | Default 1 (`john.s`); 0 = whole last name. |
| Name Prefixes to Drop | Small Text | — | e.g. `al`, `el` (only as a separate word or hyphenated). |
| Save Email In Employee Field | Data | — | Default `company_email`. |
| Create ERPNext User for the New Email | Check | — | For Microsoft sign-in to ERPNext (Social Login "Office 365"). |
| Usage Location | Data | Yes | Two-letter country code; needed before licenses can be assigned. |
| Manager Field | Data | — | Default `reports_to`. |
| Update Microsoft 365 When the Employee Changes | Check | — | Default on. |
| Field Mapping | Table | — | Employee field or path → Microsoft 365 property. Defaults: designation→jobTitle, department, cell_number→mobilePhone, name→employeeId, company→companyName. |

**Outlook Settings** — Offboarding

| Field | Type | Mandatory | Notes |
| ----- | ---- | --------- | ----- |
| Disable When | Code | — | Default `doc.status in ("Inactive", "Left")`. |
| Remove Licenses / Remove from Groups / Sign Out of All Sessions | Check | — | All default on. Distribution lists, mail-enabled security and dynamic groups are left alone (Graph can't change them). |
| Enable Again When the Create Condition Matches | Check | — | Rehire. |
| Delete Disabled Accounts After (Days) | Int | — | 0 = never. |

**Outlook Settings** — Welcome Email

| Field | Type | Mandatory | Notes |
| ----- | ---- | --------- | ----- |
| Send Welcome Email with the Password | Check | — | Default on. |
| Send To Employee Field | Data | — | Default `personal_email`. |
| Send Through | Select | — | **Microsoft 365** (Graph, from the sender mailbox; nothing stored) or **ERPNext Email** (site email account; the queued copy is deleted after sending). |
| Sender Mailbox | Data | Yes for Microsoft 365 | e.g. `it@example.com`. |
| Email Template | Link | — | Jinja with `doc`, `email`, `password`, `login_url`. |

**Microsoft 365 Provisioning Rule** — Rule Name, Enabled, Condition (Python on
the Employee, empty = everyone), Licenses, Groups. Every matching rule applies.

**Microsoft 365 Group Rule** (one per DocType — a group per document)

| Field | Type | Mandatory | Notes |
| ----- | ---- | --------- | ----- |
| DocType | Link | Yes | e.g. Project, Email Group, Opportunity. |
| Create When | Code | — | Python on the document; empty = every document. |
| Group Type | Select | Yes | **Microsoft 365** (mailbox, calendar, Teams-ready) or **Security** (access to SharePoint/apps). Needs `Group.ReadWrite.All`. |
| Group Name / Description | Data / Small Text (Jinja) | Yes / — | e.g. `PRJ - {{ doc.project_name }}`. The mail nickname is made from the name and kept unique. |
| Visibility | Select | — | Private / Public (Microsoft 365 groups). |
| Member Fields / Owner Fields | Small Text | — | Field paths, one per line: `users.user`, `owner`, `employee`, `reports_to`… Values may be Users, Employees or emails. |
| Members from Linked Documents | Table | — | Source DocType, Link Field (pointing to this DocType), Member Field, Only When (Python on the source). Email Group: `Email Group Member` / `email_group` / `email` / `not doc.unsubscribed`. |
| Remove Members No Longer Listed | Check | — | Default on. Owners are never removed automatically. |
| Close When / Then | Code / Select | — | e.g. `doc.status in ("Completed", "Cancelled")` → Remove Members, Delete Group or Nothing. |
| When the Document Is Deleted | Select | — | Keep Group (default) or Delete Group. |
| Save Group ID / Email In Field | Data | — | Optional fields on the document (e.g. custom fields) for other apps. |

People without a Microsoft 365 user (e.g. outside subscribers of an Email Group)
are listed on the **Microsoft 365 Document Group** as "Not in Microsoft 365";
guests are not invited.

### Site Config / Environment Variables
None required. `microsoft_graph_base_url` in `site_config.json` replaces
`https://graph.microsoft.com/v1.0` (local fakes only).

### Webhooks (if applicable)
None.

### How to Test
1. **Without Microsoft:** **Microsoft 365 Settings** → Enabled + **Test Mode** → Save →
   **Test Connection** says "Connected to Test Mode Tenant". **Outlook Settings** →
   Enabled, Email Domain `example.com`, Usage Location `SA` → Save →
   **Sync Licenses, Groups and Accounts**.
2. Add a **Microsoft 365 Provisioning Rule** (empty condition) with a license and
   the "All Staff" group.
3. Create an Active Employee with a personal email → a **Microsoft 365 Account**
   (Active, `first.l@example.com`) appears with the license and group; the
   Employee form shows **Microsoft 365: Active**.
4. Set the Employee to **Left** → the account becomes Disabled with no licenses or groups.
5. **Live:** fill Tenant/Client ID and secret, untick Test Mode, **Test
   Connection**, sync, then repeat 2–4 with a test employee.
6. **Group per document:** add a **Microsoft 365 Group Rule** for Project
   (Member Fields `users.user`) → create a Project with users → the form shows
   **Microsoft 365 → Open Group**; the **Microsoft 365 Document Group** lists the members.
7. Problems: **Microsoft 365 Log** (Microsoft's error message) and **Last Error**
   on the account.

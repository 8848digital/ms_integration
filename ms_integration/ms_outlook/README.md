<!--
Copyright (c) 2026 8848 Digital LLP. All rights reserved.
Proprietary and confidential. Unauthorized copying, distribution, or use
of this file, via any medium, is strictly prohibited without prior
written permission from 8848 Digital LLP.
-->

# MS Outlook

## Purpose
Outlook / Microsoft 365 accounts from ERPNext: account creation (onboarding),
updates, offboarding and rehire, provisioning rules, license/group editing,
the directory sync and a Microsoft 365 group per document. Uses the connection,
Graph client and log from **MS Core**.

## DocTypes
| DocType | Purpose |
| ------- | ------- |
| Outlook Settings | Single: Enabled, address pattern, conditions, offboarding and welcome-email options, field mapping, directory sync. |
| Microsoft 365 Field Map | Child of Outlook Settings: Employee field → Microsoft 365 user property. |
| Microsoft 365 Provisioning Rule | Condition on the Employee → licenses and groups to give. |
| Microsoft 365 License | License types of the tenant with purchased / assigned / available seats (synced). |
| Microsoft 365 Group | Groups of the tenant with type and whether members can be changed (synced). |
| Microsoft 365 License Item / Group Item | Children (Table MultiSelect) of rules and accounts. |
| Microsoft 365 Account | One employee's account: status, sign-in name, licenses and groups (editing rows changes Microsoft 365). |
| Microsoft 365 Group Rule | One per DocType: when to create a group per document, its type and name, members/owners from fields and linked documents, what happens on close/delete. |
| Microsoft 365 Group Member Source | Child of Group Rule: documents linking to the rule's document that add members (e.g. Email Group Member). |
| Microsoft 365 Document Group | The group of one document: status, group id/email, members and whether each was added. |
| Microsoft 365 Document Group Member | Child of Document Group: one person, role, status. |

## Customizations
| DocType | Owned by | What's customized |
| ------- | -------- | ------------------ |
| All DocTypes (`"*"`) | Any | `group_events.py`: on_update / on_update_after_submit queue the document's group sync (and of documents fed by a member source); on_trash drops the group record (and group if configured); after_delete resyncs groups fed by a deleted source. Cached early exit for DocTypes without rules. |
| Forms of DocTypes with a Group Rule | Any | bundle JS: **Microsoft 365 → Open Group / Create Group / Sync Group Members**. |
| Employee | ERPNext | `on_update` queues create / update / disable / enable of the account (`customization/employee/`); form shows the account status with **Create Account / Open Account** buttons. |

## Code layout
- `api/v1/account.py` — `get_status`, `create_account`, `run_action` (disable, enable, reset_password, refresh).
- `api/v1/directory.py` — `sync_directory`.
- `api/v1/group.py` — `get_group`, `sync_group`.
- `group_events.py` — doc_events "*" for Group Rules.
- `services/` — `settings` (Outlook Settings, enabled check), `directory_api` + `account_api` + `groups_api` (Graph operations), `addresses`, `payload`, `welcome`, `access` (rules, row edits), `provisioning` + `provisioning_steps`, `lifecycle`, `directory` (sync), `people` (addresses → user ids), `group_rules`, `document_groups` + `document_group_close`.
- `tasks.py` — background jobs (`provision_account`, `update_account`, `disable_account`, `enable_account`, `sync_directory`, `sync_document_group`, `delete_group`) and scheduled `daily_sync` (also re-checks active document groups), `delete_expired_accounts`.

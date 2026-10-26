<!--
Copyright (c) 2026 8848 Digital LLP. All rights reserved.
Proprietary and confidential. Unauthorized copying, distribution, or use
of this file, via any medium, is strictly prohibited without prior
written permission from 8848 Digital LLP.
-->

# MS Core

## Purpose
What every Microsoft 365 module of this app shares: the connection to the
tenant (app registration credentials, Test Mode), the Microsoft Graph client,
the Test Mode fake tenant and the Microsoft 365 Log.

## DocTypes
| DocType | Purpose |
| ------- | ------- |
| Microsoft 365 Settings | Single: Enabled, Test Mode, timeout, tenant ID, client ID, client secret. |
| Microsoft 365 Log | Every change made in Microsoft 365 and every failure, from any module. |

## Customizations
None.

## Code layout
- `api/v1/settings.py` — `test_connection` (tenant name, verified domains).
- `services/graph.py` — Graph client: cached client-credentials token, retries on throttling, paging, errors.
- `services/test_mode.py` + `test_mode_routes.py` + `test_mode_group_routes.py` — fake tenant (JSON file in the site's private folder) answering the Graph calls in Test Mode.
- `services/logs.py` — Microsoft 365 Log entries.
- `services/settings.py` — connection settings, rule conditions, field values.
- `services/field_paths.py` — field paths through child tables and links.

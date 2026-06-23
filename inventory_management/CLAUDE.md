# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the App

```bash
cd inventory_management
python3 app.py          # starts Flask on http://localhost:5001
```

Flask must be run from the `inventory_management/` directory. The `scripts/` directory is added to `sys.path` at startup, so all imports work relative to that. Default admin credentials: `admin` / `admin` (seeded on first run if no users exist).

## Legacy Excel Scripts

The scripts in `scripts/` predate the Flask app and operate directly on `inventory_master.xlsx`:

```bash
python3 scripts/parse_manifest.py        # parse a container manifest CSV → updates Excel
python3 scripts/sort_inventory.py        # re-sort design blocks alphabetically
python3 scripts/refresh_counts.py        # recalculate QTY counts and variance
python3 scripts/apply_status_fills.py    # reapply status colors + column widths
```

Dependency: `openpyxl`. These scripts are largely superseded by the Flask app but kept for bulk Excel operations.

## Architecture

### Two parallel systems
The repo started as pure Excel automation and evolved into a Flask web app. Both still exist:
- **Excel layer** (`scripts/` + `inventory_master.xlsx`) — legacy, for bulk ops
- **Web layer** (`app.py` + `inventory.db`) — primary, used day-to-day

The Flask app calls `export_excel(conn)` after mutating the DB so the Excel file stays in sync.

### Database (`scripts/db.py`)
Single SQLite file at `inventory.db`. All DB access goes through `db.py` — `app.py` imports everything from there. Schema tables:

| Table | Purpose |
|---|---|
| `units` | Every physical unit; statuses: `In Stock`, `Pre-Sale (CNT-XXX)`, `In Production`, `Allocated (ORD-XXX)` |
| `variants` | Unique (design, size, finish, swing, glass) combos with SKU and optimal count |
| `sales_orders` | Active pre-arrival allocations |
| `warehouse` | Units that have physically arrived and are being prepped |
| `warehouse_checklist` | 14-item QC checklist per warehouse unit (keyed on `warehouse_id`) |
| `warehouse_photos` | Photos per warehouse unit |
| `users` | Roles: `admin` or `warehouse` |
| `activity_log` | Append-only audit trail of all user actions |

`init_db()` creates all tables and runs pending `ALTER TABLE` migrations (wrapped in try/except). New columns must be added both to the `SCHEMA` string and to the migrations list.

### Unit lifecycle
```
In Production → Pre-Sale (CNT-XXX) → [allocated → sales_orders] → warehouse → Pending Review → Ready for Pickup / Ready for Delivery
```
- Units with `status LIKE 'Allocated%'` are filtered out of the inventory page.
- `_restore_status(conn, serial, container_id)` determines whether a cancelled unit returns to `In Stock` or `Pre-Sale` by checking `units.date_received`.
- `move_to_warehouse()` copies a `sales_orders` row to `warehouse` (carrying `fulfillment_type`) and deletes the source row.

### Warehouse prep workflow
Each `warehouse` row gets its own prep sheet at `/warehouse/prep/<wh_id>`. The 14 checklist items are defined in `CHECKLIST_ITEMS` at the top of `db.py`. When all 14 are checked, the "Mark Ready" button appears → sets status to `Pending Review`. Office staff then Accept (→ `Ready for Pickup` or `Ready for Delivery` based on `fulfillment_type`) or Push Back (→ `In Prep`).

### Auth and roles
`login_required` — any logged-in user. `admin_required` — admin only. A context processor injects `current_user` and `is_admin` into every template. Warehouse users are redirected to `/warehouse` on login and cannot see other tabs or cancel/change orders. Use `generate_password_hash(password, method='pbkdf2:sha256')` — `scrypt` is not available on this system's OpenSSL.

### Fulfillment type
`fulfillment_type` (`pickup` / `delivery`) lives on both `sales_orders` and `warehouse`. It is currently randomly assigned (Shopify will set it in the future). The warehouse tab is split into four sections by aggregate order status, and each section shows the fulfillment badge.

### Templates
All templates extend `base.html`. Grouped order views (sales, warehouse) use a Jinja `unit_map` dict built via `namespace` to map `order_number → [units]`. Expand/collapse is handled by `toggleOrder()` in inline `<script>` blocks. AJAX calls (checklist, notes) post to `/warehouse/prep/<wh_id>/<action>` and return `{"ok": true}`.

### Key conventions
- `scripts/constants.py` is the single source of truth for file paths, column indices, and the SKU map. Import `_BASE_DIR` from there when you need to resolve paths relative to the project root.
- Serial format: `YY-MMDD-####` (year ordered, date MMDD, sequential number resetting each year).
- SKU format: `DESIGN-WxH-FINISH-SWING-GLASS` (e.g. `VAL-36X80-MB-RI-CLR`). `get_sku()` in `constants.py` looks up the curated map first, then falls back to dynamic generation.
- Photo uploads stored at `static/uploads/<wh_id>/<timestamped_filename>`.

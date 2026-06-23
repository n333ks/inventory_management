# Inventory Management System

Automated inventory tracking for an iron doors business. Manages the full lifecycle of a unit — from production order through container arrival to in-stock — inside a single Excel workbook (`inventory_master.xlsx`).

---

## Folder Structure

```
inventory_management/
├── inventory_master.xlsx       ← master inventory workbook
├── README.md
├── receiving_log.txt
├── manifests/                  ← one CSV per container
│   └── container_manifest_CNT-XXX.csv
├── purchase_orders/            ← generated POs (po_YYYY-MM-DD.csv)
└── scripts/                    ← all Python scripts
    ├── constants.py            ← single source of truth for paths, columns, SKUs
    ├── parse_manifest.py
    ├── apply_status_fills.py
    ├── refresh_counts.py
    ├── sort_inventory.py
    ├── seed_production_rows.py
    └── fix_inventory_formatting.py
```

---

## Workflow Overview

```
Units ordered → In Production rows added → Container manifest parsed → Pre-Sale (CNT-XXX) → Container arrives → In Stock
```

1. **Negative variance detected** — run `generate_po.py` to create a PO and seed In Production rows simultaneously
2. **Container manifest created** — saved to `manifests/container_manifest_CNT-XXX.csv`
3. **Container departs** — parse with `parse_manifest.py` → serials filled in, status flips to `Pre-Sale - CNT-XXX` (yellow) or `Allocated - ORD-XXX` (blue) for customer orders
4. **Container arrives** — manually update status to `In Stock` (green)

---

## Inventory Master (`inventory_master.xlsx`)

Each row represents one physical unit. Columns:

| Column | Field | Notes |
|--------|-------|-------|
| A | Design Name | e.g. Valencia Single Door |
| B | Size | e.g. 36" x 80" |
| C | Finish | e.g. Matte Black |
| D | Swing | e.g. Left Inswing |
| E | Glass Type | e.g. Clear |
| F | SKU | Auto-populated from `constants.py` SKU map |
| G | In-Stock QTY | Auto-calculated |
| H | In-Production QTY | Auto-calculated |
| I | Pre-Sale QTY | Auto-calculated |
| J | Optimal Count | Set manually — target stock level |
| K | Variance | (G+H+I) − J · Red if negative = reorder needed |
| L | Serial Number | Format: `YY-MMDD-####` |
| M | Status | In Stock / Pre-Sale - CNT-XXX / In Production / Allocated - ORD-XXX |

**Serial format:** `YY-MMDD-####` — year ordered, date ordered (MMDD), sequential number resetting each calendar year. Example: `26-0502-0010` = ordered May 2nd 2026, unit #10.

**Row spacing:** 1 blank row between variants of the same design, 2 blank rows between designs. Designs are sorted alphabetically. The first row of each variant shows counts (G–K) in bold; repeated label rows below use white font to stay visually clean. Column widths auto-fit to content on every script run.

**Status colors:**
- 🟢 Green — In Stock
- 🟡 Yellow — Pre-Sale (on a container, no customer assigned)
- 🔴 Red — In Production (no serial yet)
- 🔵 Blue — Allocated (reserved for a specific customer order)

---

## SKUs

Every variant has a SKU defined in `scripts/constants.py` (`SKU_MAP`). Format:

```
DESIGN-SIZExSIZE-FINISH-SWING-GLASS
```

Example: `VAL-36X80-MB-RI-CLR` = Valencia Single Door, 36×80, Matte Black, Right Inswing, Clear.

SKUs are written automatically to column F in the inventory sheet and to the `SKU` column in every container manifest. When adding a new design, add one entry to `SKU_MAP` in `constants.py` and all scripts pick it up automatically.

**Design codes:** ALT · CAD · COR · GRN · MAL · RON · SEV · TOL · VAL

---

## Scripts

All scripts live in `scripts/` and can be run from any directory — paths are resolved automatically via `constants.py`.

### Day-to-day

| Script | Purpose | Usage |
|--------|---------|-------|
| `parse_manifest.py` | Parse a container manifest and update inventory | `python3 scripts/parse_manifest.py` |
| `sort_inventory.py` | Re-sort all design blocks alphabetically | `python3 scripts/sort_inventory.py` |
| `seed_production_rows.py` | Add In Production rows for upcoming containers | Edit `SEED` dict, then run |
| `apply_status_fills.py` | Re-apply all status colors and column widths | `python3 scripts/apply_status_fills.py` |
| `refresh_counts.py` | Recalculate QTY counts and variance | `python3 scripts/refresh_counts.py` |

### Supporting modules (imported by the above — do not run directly)

| Script | Purpose |
|--------|---------|
| `constants.py` | All column indices, file paths, SKU map, and `get_sku()` helper |
| `fix_inventory_formatting.py` | Border formatting; exports `apply_design_block_borders()` |

### Legacy / one-time scripts
Used during initial setup and kept for reference only.

`add_row_spacing.py`, `container_receiving_tool.py`, `create_inventory_master.py`, `fix_detail_rows.py`, `fix_summary_rows.py`, `inventory_update.py`, `rebuild_inventory.py`, `seed_initial_inventory.py`, `test_presale_rows.py`

---

## Container Manifests

Stored in `manifests/`. Named `container_manifest_CNT-XXX.csv`. Columns:

```
Design Name, Size, Finish, Swing, Glass Type, SKU, Quantity, Serial Numbers, Customer Order
```

- Rows sorted alphabetically by Design Name
- Serial numbers separated by `;`
- `Customer Order` column is blank for standard Pre-Sale units; populated (e.g. `ORD-1001 - Martinez Residence`) for units allocated to a specific customer
- Run `python3 scripts/parse_manifest.py` and select the container at the interactive prompt

---

## Purchase Orders

Stored in `purchase_orders/`. Named `po_YYYY-MM-DD.csv`.

Generated automatically when negative variance is detected — the PO lists every variant that has fallen below its optimal count, with the exact quantity needed to return to zero variance. In Production rows are seeded in the inventory master simultaneously.

---

## Parse Logic

- **Matched variant** (exists in inventory):
  - Fills blank In Production rows with serials → status becomes `Pre-Sale - CNT-XXX`
  - Units with a `Customer Order` value → status becomes `Allocated - ORD-XXX`
  - Any serials beyond existing In Production rows are appended directly to the variant block (no duplicate line created)
- **New variant / new design**:
  - Inserted at the correct alphabetical position
  - 1 blank row separator between variants of the same design
  - 2 blank rows between different designs

---

## Running the Scripts

```bash
cd ~/Desktop/learning-log/inventory_management
python3 scripts/parse_manifest.py        # interactive manifest selector
python3 scripts/sort_inventory.py        # re-sort alphabetically
python3 scripts/refresh_counts.py        # recalculate all counts
python3 scripts/apply_status_fills.py    # reapply colors + column widths
```

Dependencies: `openpyxl` (`pip install openpyxl`)

---

## Future: Zoho / CRM Integration

The SKU column is the bridge to any inventory API (Zoho Inventory, etc.). When transitioning away from Excel, the parse script would call the Zoho API instead of writing to the workbook — looking up each SKU to update quantities and statuses. Container manifests are plain CSV, making them universally compatible with any import tool.

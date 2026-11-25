# TE Automation — Brazil

> **Context file for the TE-Automation-Brazil repo.** Auto-loaded by Claude Code sessions. Keep it current.

---

## 1. What this is

Desktop application that automatically fills **True Efficiency (TE) reporting templates** for PepsiCo Brazil's production plants. It reads the monthly True Efficiency exports (a folder of four Power BI `.xlsx` files — see §4), extracts the per-plant/per-line numbers, and writes them into a `_Filled` copy of each plant's blank template — in minutes, instead of the week-plus it takes by hand.

This repo is the **Brazil instance** of the TE Automation initiative. It is a fork of the Mexico POC:

- **Reference implementation:** [`TE-Automation` (Mexico)](https://github.com/p3droJosue/TE-Automation)
- **Mexico status:** in production, 15 plants, ~35 min per cycle
- **Brazil scope:** 8 template files / 43 line blocks (Amacoco Petrolina, CIC, Feira de Santana, ITU, Lucky, NE Recife, Sete Lagoas, Sorocaba)

---

## 2. Why a separate repo (strategic choice)

We deliberately did **not** merge Brazil into the Mexico app. Reasons:

- Each country ships on its own cadence. Brazil doesn't wait for Mexico and vice-versa.
- Analysts in Brazil see a plant list they can scan — not a cross-country menu with 200+ entries.
- Blast radius of a `plants.py` change stays within one country.
- Country ownership is clean: the Brazil team owns the Brazil repo.

**Long-term:** once Brazil is stable, the plan is to extract `engine.py` + the UI shell into a shared `te_automation_core` package that both Mexico and Brazil (and later APAC/AMESA/EMEA) import. Each country repo then owns only its `plants.py` + branding. See §8 — Future consolidation.

---

## 3. Architecture (same shape as Mexico)

Three modules, one responsibility each:

| File                     | Owns                                                        |
| ------------------------ | ----------------------------------------------------------- |
| `app.py`                 | UI (CustomTkinter). Tabs, sidebar, status pills, threading. |
| `engine.py`              | Reads the DB exports, plans cell writes, applies them in Excel (xlwings). |
| `plants.py`              | Data registry — plants, line blocks, DB line names, SKU master. |

Dependency direction:

```
app.py  ─▶  engine.py  ─▶  (openpyxl to read exports, xlwings to write templates)
app.py  ─▶  plants.py  ─▶  engine.py (uses Job / PlantSpec)
```

**Brazil diverged from Mexico in `engine.py`'s reading/writing half** (different DB format — see §4). What is still shared with Mexico and should keep flowing both ways:

- OneDrive staging (copy to temp before processing — avoids `"Cannot access"` COM errors)
- Excel cleanup in `_fill_plant_core`'s `finally` (always quit Excel)
- mtime verification (treat as success if the output file was actually rewritten, even if COM throws a spurious error after)
- Fixed-height scrollable plant list with theme button pinned to the top of the sidebar
- Apple-style per-plant status pills (Pending / Running / Done / Failed) with live summary
- Clear plain-English error translation via `_short_error()`
- Build config that bundles the logo asset into a single-file `.exe`

If you find yourself editing `engine.py` or `app.py`, pause and ask: *is this a Brazil-specific need, or a general improvement that Mexico should also get?* General improvements should be ported back.

---

## 4. How the Brazil engine works

### 4.1 Inputs

- **Databases folder** (`databases/`) with four Power BI exports, recognised by content (file names don't matter):
  - APP downtime pivot (row 2 label `TEMPO_DBL`) — most lines
  - ITA downtime pivot (`HORAS ITA`) — Lucky
  - MES downtime pivot (`DT Horas`) — Sete Lagoas + ITU PC50 / DORITOS_TC1 / CLEXTRAL I & II
  - Production list (`Planta | Linha | Corte | Sabor | Produção Real KG`)
- **Blank template** `Snk_TE_<Plant>_P11 2026 OCT(P10 UK).xlsm`. The month in the name sets *Total calendar time* (days × 24).

### 4.2 Pipeline (`engine.py`)

1. `load_brazil_db(folder)` → `BrazilDB` (downtime per source/line/category, production per line).
2. `plan_plant(db, plant, calendar_hours)` → list of `Write(sheet, row, col, value)` + warnings. **Pure Python, no Excel**, so it is testable anywhere.
3. `fill_plant(...)` opens the blank in Excel (xlwings), applies the writes, recalculates and saves to `<template>_Filled.xlsm`. Returns the warnings; the app also saves them to `<template>_Filled_warnings.txt`.

Key tables in `engine.py`:

- `CATEGORY_CELLS` — DB loss category → template cell (Default layout; TC rows +1, PC rows +10). Includes the MES coarse buckets and their **assumed** destinations (pending Brazil team confirmation, see `analysis/BRAZIL_ANALYSIS.md` §5).
- `PRODUCT_LAYOUT` / `PRODUCT_COLS` — product & production tables per layout. Column G is a formula on PC/TC/Stax and is never written.

Hard rules (enforced by `analysis/check_against_golden.py`, must report `SAFETY=0`):

- never write to a cell that holds a formula in the blank template;
- never write inside a merged area (Excel refuses);
- only write real values (output always starts from a fresh copy of the blank).

### 4.3 `plants.py` — the registry + SKU master

```python
Job("Extruded", 0, "4 - Clextral", "APP", "CLEXTRAL III", type_value="Baked", te_target=0.785444, skus=(
    SKU("CEBOLITOS", ptype="MD", seasoning=0.3331, rated=960.336, db_keys=(("CEBOLITOS", "CEBOLLA"),)),
    ...
)),
```

- `sheet`, `cons` — template sheet and block index (0 = columns B..I).
- `header_text` — line name exactly as in the template's dropdown.
- `source` / `db_line` — which downtime export, and the `Linha` value in the exports.
- `type_value`, `te_target`, `fryer` — static per-line values copied from the goldens.
- `skus` — product table top to bottom. `db_keys` lists the `(Corte, Sabor)` pairs whose kg feed that row. A production row with no matching SKU is appended as a new row and reported as a warning ("not in the SKU master") — the fix is to add the key (or a new `SKU`) here.

`plants.py` was bootstrapped from the OCT 2026 goldens by `analysis/build_plants_from_goldens.py`; it is now edited by hand.

### 4.4 Not filled by the engine (analyst fills by hand)

Rate/T-put losses (G106–G111), PC gross solids (C17), # FG SKUs, Stax line type (formula). Anything the DB doesn't carry for that month stays blank.

### 4.5 What you should NOT change

- OneDrive staging (`fill_plant`, `_is_onedrive_path`) and the mtime safety net.
- UI structure in `app.py` — Single Plant tab, Run All tab, status cards, status pills.
- The safety rules in §4.2.

---

## 5. Testing & running

### 5.1 Without Excel (any OS) — check the plan against the goldens

```bash
pip install openpyxl
python analysis/check_against_golden.py            # summary per plant; SAFETY must be 0
python analysis/check_against_golden.py ITU -v     # every DIFF / MANUAL cell for one plant
```

`DIFF` today is expected: the OCT exports exclude 27–31 Oct while the goldens cover the full month, and MES lacks the dark-hour / breakdown detail (see `analysis/BRAZIL_ANALYSIS.md`).

### 5.2 With Excel (Windows) — the real run

```powershell
git clone https://github.com/p3droJosue/TE-Automation-Brazil.git
cd TE-Automation-Brazil
git checkout claude/eager-galileo-fbakkz
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install customtkinter xlwings openpyxl pillow pyinstaller
cd modular_architecture
python app.py
```

- **Run All tab:** Databases folder = `..\databases`, Templates folder = `..\templates\TE_blank_templates` → one `_Filled.xlsm` per plant next to the blanks (blanks are never overwritten).
- **Single Plant tab:** pick a plant in the sidebar, Databases folder, and one blank template.
- **CLI:** `python app.py --db ..\databases --folder ..\templates\TE_blank_templates --all`
- Then open each `_Filled.xlsm` next to the golden in `templates\TE_completed_templates\` and compare.

### 5.3 Build the .exe

```powershell
pyinstaller build.spec      # → dist\TE_Filler_Brazil.exe
```

---

## 6. Keeping in sync with Mexico fixes

Since Brazil and Mexico are forks, bug fixes don't flow automatically. The workflow:

- **New fix lands on Mexico:** open its PR on `p3droJosue/TE-Automation`.
- **Mirror in Brazil:** fixes to `app.py` (UI) and to the staging / Excel-cleanup parts of `engine.py` port directly. Fixes to Mexico's RDP/Pond readers and writers do **not** apply — Brazil has its own.
- **Test in Brazil:** quick Run All to confirm no Brazilian regression.
- **Push.**

Once the shared-core refactor lands (§8), this manual sync goes away.

---

## 7. Common pitfalls for Brazil specifically

- **Export period.** The exports must cover the same days as the template (the OCT 2026 files exclude 27–31 Oct, so they don't reproduce the goldens).
- **MES coarse buckets.** MES has no dark-hour or Tech/Op split; `CATEGORY_CELLS` puts them on assumed cells (Other / Technical / External). Confirm with the Brazil team.
- **New products.** A new Corte/Sabor in the production export shows up as a warning and an extra product row without type/throughput (Excel shows #DIV/0 on that row) — add it to the SKU master in `plants.py`.
- **Filename pattern.** `match_filename()` matches plant keywords in `Snk_TE_<Plant>_P11 2026 OCT(P10 UK).xlsm`; the month/year in the name drives calendar hours.
- **OneDrive tenant naming.** The staging detector is case-insensitive and matches any path containing "onedrive".

### 7.1 What we actually found (OCT 2026 analysis)

See `analysis/BRAZIL_ANALYSIS.md` for the full write-up, and run `python analysis/compare_db_vs_golden.py` to re-check DB vs golden.

- **The template is identical to Mexico's** (same sheets/blocks/rows). Brazil also uses the `Stax` sheet (Sorocaba Canister).
- **The DBs are NOT Mexico-shaped.** There is no `Reporte Detalle Plantas` / `Pond`. Instead there are 4 Power BI exports in `databases/`: downtime pivots per source system (**APP**, **ITA** = Lucky, **MES** = Sete Lagoas + some ITU lines) and one production list (Planta/Linha/Corte/Sabor/kg). Category names are already English TE categories. The *reader* half of `engine.py` must be replaced; the writer/staging half stays.
- Product master data (template SKU name, type, % seasoning, rated throughput) is not in any DB, so it needs a per-line SKU master.
- Templates live in `templates/TE_blank_templates/` and `templates/TE_completed_templates/` (goldens, suffix `_complete`).

---

## 8. Future consolidation (the shared-core plan)

Once Brazil is working and we've learned what truly differs between countries, extract the common code:

```
te-automation/
├── te_core/                      # shared Python package (pip-installable)
│   ├── engine.py                 # fill engine, staging, mtime check
│   ├── ui.py                     # CustomTkinter widgets + App shell
│   ├── writers.py                # _write_mandatory, _write_pc, etc.
│   └── build_spec_template.py    # PyInstaller template
│
├── te-automation-mexico/
│   ├── plants.py                 # Mexican registry
│   ├── app.py                    # 20 lines — imports te_core.App, sets branding
│   └── build.spec
│
└── te-automation-brazil/
    ├── plants.py                 # Brazilian registry
    ├── app.py                    # 20 lines — imports te_core.App, sets branding
    └── build.spec
```

Trigger to start the refactor: when APAC or AMESA comes on board and we're staring down a third fork.

---

## 9. Quick reference

| Need to…                                   | Edit                                       |
| ------------------------------------------ | ------------------------------------------ |
| Add a line / plant                         | `plants.py` — new `Job` / `PlantSpec` + `REGISTRY` + `_FNAME_KEYWORDS` |
| Map a new product (Corte/Sabor)            | `plants.py` — `SKU(..., db_keys=...)`      |
| Fix a wrong hour cell                      | `engine.py` — `CATEGORY_CELLS`             |
| Check the plan without Excel               | `python analysis/check_against_golden.py`  |
| Change window title                        | `app.py` — `self.title(...)`               |
| Change `.exe` name                         | `build.spec` — `name="..."`                |
| Rebuild EXE                                | `pyinstaller build.spec` (close old .exe first) |
| See where the EXE lands                    | `modular_architecture/dist/TE_Filler_Brazil.exe` |

| Who owns what                              |                                            |
| ------------------------------------------ | ------------------------------------------ |
| Repository                                 | Brazil TE lead                             |
| Mexico reference                           | Pedro Josue (p3droJosue)                   |
| Engine improvements                        | Shared — port both directions              |
| Brazilian plant registry                   | Brazil analyst team                        |

---

## 10. End-user requirements (unchanged from Mexico)

- Windows 10 or 11 (64-bit)
- Microsoft Excel 2016+ installed
- ~150 MB disk for the .exe
- No Python, no pip, no install

The .exe is distributed via OneDrive / Teams / USB. On first run, Windows SmartScreen shows a blue "Windows protected your PC" dialog — click *More info* → *Run anyway*.

# TE Automation — Brazil

> **Context file for the TE-Automation-Brazil repo.** Auto-loaded by Claude Code sessions. Keep it current.

---

## 1. What this is

Desktop application that automatically fills **True Efficiency (TE) reporting templates** for PepsiCo Brazil's production plants. It reads a monthly consolidation database (`.xlsm`), extracts the per-plant/per-line numbers, and writes them into each plant's blank template — in minutes, instead of the week-plus it takes by hand.

This repo is the **Brazil instance** of the TE Automation initiative. It is a fork of the Mexico POC:

- **Reference implementation:** [`TE-Automation` (Mexico)](https://github.com/p3droJosue/TE-Automation)
- **Mexico status:** in production, 15 plants, ~35 min per cycle
- **Brazil goal:** same outcome for Brazilian plants (TBD: confirm plant count with the Brazil team)

---

## 2. Why a separate repo (strategic choice)

We deliberately did **not** merge Brazil into the Mexico app. Reasons:

- Each country ships on its own cadence. Brazil doesn't wait for Mexico and vice-versa.
- Analysts in Brazil see a plant list they can scan — not a cross-country menu with 200+ entries.
- Blast radius of a `plants.py` change stays within one country.
- Country ownership is clean: the Brazil team owns the Brazil repo.

**Long-term:** once Brazil is stable, the plan is to extract `engine.py` + the UI shell into a shared `te_automation_core` package that both Mexico and Brazil (and later APAC/AMESA/EMEA) import. Each country repo then owns only its `plants.py` + branding. See §8 — Future consolidation.

---

## 3. Architecture (identical to Mexico)

Three modules, one responsibility each:

| File                     | Owns                                                        |
| ------------------------ | ----------------------------------------------------------- |
| `app.py`                 | UI (CustomTkinter). Tabs, sidebar, status pills, threading. |
| `engine.py`              | Excel automation (xlwings). All business logic.             |
| `plants.py`              | Data registry — which plants, which Jobs, which RDP rows.   |

Dependency direction:

```
app.py  ─▶  engine.py  ─▶  (xlwings, pandas)
app.py  ─▶  plants.py  ─▶  engine.py (uses Job / PlantSpec)
```

**Country-specific code lives in `plants.py`.** Everything else should stay identical to Mexico's version, which gets you the bug fixes "for free":

- OneDrive staging (copy to temp before processing — avoids `"Cannot access"` COM errors)
- mtime verification (treat as success if the output file was actually rewritten, even if COM throws a spurious error after)
- Fixed-height scrollable plant list with theme button pinned to the top of the sidebar
- Apple-style per-plant status pills (Pending / Running / Done / Failed) with live summary
- Clear plain-English error translation via `_short_error()`
- Build config that bundles the logo asset into a single-file `.exe`

If you find yourself editing `engine.py` or `app.py`, pause and ask: *is this a Brazil-specific need, or a general improvement that Mexico should also get?* General improvements should be ported back.

---

## 4. What to change for Brazil

### 4.1 `plants.py` — the plant registry (95% of the work)

This is where Brazil differs from Mexico. Replace the full file with Brazilian plants:

```python
from engine import Job, PlantSpec

SAO_PAULO = PlantSpec(
    key="SAO_PAULO",
    jobs=[
        Job("Extruded", "DEFAULT", 0, "1 - Line A1", 15, "PLANTA SAO PAULO", "A1", type_value="Fried"),
        Job("Extruded", "DEFAULT", 1, "2 - Line A2", 16, "PLANTA SAO PAULO", "A2", type_value="Baked"),
        # ... one Job per line/product on each product sheet
    ]
)

# Repeat for each Brazilian plant

REGISTRY = {
    "SAO_PAULO": SAO_PAULO,
    # ...
}

def match_filename(fname: str) -> str | None:
    """Map a Brazilian template filename to a REGISTRY key."""
    # e.g. "Snk_TE_Sao Paulo_P03 2026 ..." -> "SAO_PAULO"
    ...
```

**Per-Job rules (same as Mexico):**

- `sheet` — one of `"PC"`, `"TC"`, `"Other"`, `"Extruded"`, `"Cookies or Biscuits"`, `"Wafer"`. The Brazilian template should use the same sheet names; if it doesn't, update `engine._PRODUCT_SHEETS` too.
- `kind` — `"DEFAULT"`, `"TC"`, or `"PC"`. Picks the writer function.
- `cons` — block index (0, 1, 2, …) within the sheet. Must be contiguous per sheet.
- `rdp_row` — the row number in `Reporte Detalle Plantas` (Brazilian DB) for this line.
- `pond_plant` / `pond_line` — the plant name and line code as they appear in the `Pond` sheet.

### 4.2 `engine.py` constants — only if Brazil needs them

Review and adjust if the Brazilian DB differs:

- **`MANDATORY_HEADERS`** — must appear in the DB's RDP sheet. If Brazilian DB uses Portuguese names ("Horas Totais" instead of "Total Hours"), either (a) translate the list here, or (b) ask the Brazil team if English headers can be preserved. Option (a) is fine; keep the list in Portuguese if that's the source of truth.
- **`OPTIONAL_HEADERS`** — same, in Portuguese if needed.
- **`ROWS_TEMPLATE_DEFAULT` / `_TC` / `_PC`** — only if the Brazilian template has different row layouts. Unlikely but verify by opening a sample.
- **`BLOCK_WIDTH`** — 8 columns per block. Almost certainly identical.
- **`_PRODUCT_SHEETS`** — sheet names in the template. If translated to Portuguese, update here.

### 4.3 Branding in `app.py`

Minimal changes:

```python
self.title("TE Template Filler — Brazil")         # was: "— Mexico"
# ...
ctk.CTkLabel(sb, text="Brazil Sites", ...)        # was: "Mexico Sites"
```

Logo: PepsiCo logo stays the same (global brand). If a Brazil-specific sub-brand is wanted, drop a new PNG in `assets/` and point `logo_path` at it.

### 4.4 Build output name

In `build.spec`:

```python
exe = EXE(
    ...,
    name="TE_Filler_Brazil",      # was: "TE_Filler"
    ...
)
```

So the `.exe` is distinguishable in Explorer / Task Manager from the Mexico one.

### 4.5 What you should NOT change

- Core fill loop in `engine.py` (`_fill_plant_core`) — the open/read/write/save sequence is country-agnostic.
- OneDrive staging logic (`fill_plant`, `_is_onedrive_path`) — benefits Brazil immediately.
- mtime verification — same.
- UI structure in `app.py` — Single Plant tab, Run All tab, status cards, status pills.
- Writer functions (`_write_mandatory`, `_write_optional`, `_write_pc`, `_write_tc_sku`, `_write_default_sku`) — these implement the template's cell layout, which is the same per the Brazil team.

---

## 5. Getting started (Brazil onboarding checklist)

1. **Get artifacts from the Brazil team:**
   - One monthly TE database `.xlsm` (recent month).
   - One **completed** template per plant (the "golden" to diff against).
   - One **blank** template per plant.
   - A mapping sheet: for each Brazilian plant, which lines run on which product sheet (`PC` / `TC` / `Other` / `Extruded` / `Cookies or Biscuits` / `Wafer`), and which RDP row each line corresponds to.

2. **Set up the dev environment:**
   ```powershell
   git clone <this-repo>
   cd TE-Automation-Brazil
   python -m venv .venv
   .venv\Scripts\Activate.ps1
   pip install customtkinter xlwings pandas pillow openpyxl pyinstaller
   ```

3. **Verify the Mexican code still runs** (sanity check that the fork works):
   ```powershell
   cd modular_architecture
   python app.py
   ```
   It'll still show Mexican plants — that's expected before you replace `plants.py`.

4. **Replace `plants.py`** with the Brazilian registry (see §4.1).

5. **Test one plant end-to-end:**
   - Pick one simple Brazilian plant.
   - Run the Single Plant tab.
   - Diff the filled output against the golden template.
   - Fix any `rdp_row` / `pond_line` mismatches until they match.

6. **Expand to all plants**, then run the Run All tab against the folder containing every Brazilian blank template.

7. **Build the .exe** and ship to the Brazil analyst team:
   ```powershell
   pyinstaller build.spec
   # output: dist/TE_Filler_Brazil.exe
   ```

---

## 6. Keeping in sync with Mexico fixes

Since Brazil and Mexico are forks, bug fixes don't flow automatically. The workflow:

- **New fix lands on Mexico:** open its PR on `p3droJosue/TE-Automation`.
- **Mirror in Brazil:** cherry-pick the commit (or manually apply the diff) to this repo. The diff usually touches `engine.py` or `app.py` and is identical.
- **Test in Brazil:** quick Run All to confirm no Brazilian regression.
- **Push.**

Once the shared-core refactor lands (§8), this manual sync goes away.

---

## 7. Common pitfalls for Brazil specifically

- **Portuguese column headers in the database.** If the Brazilian DB uses "Horas Totais" / "Manutenção" / etc., `MANDATORY_HEADERS` has to be translated. Check with the Brazil team which language the DB is published in.
- **Different plant/line naming in Pond.** `pond_line` values are case-insensitive but exact-match. Brazilian codes may be `L1`, `L2` instead of `A1`, `A2`. Open the Pond sheet to confirm.
- **Filename patterns.** Mexican templates follow `Snk_TE_<Plant>_P<period> <year> <month>(P<x> UK).xlsm`. Brazilian names may differ; update `match_filename()` accordingly or the Run All tab won't match anything.
- **OneDrive tenant naming.** Mexico uses `OneDrive - Pepsico`. Brazil may be `OneDrive - PepsiCo` (case) or `OneDrive - Pepsico do Brasil`. The staging detector is case-insensitive and matches any folder containing "onedrive", so this should Just Work — but confirm.
- **`Reporte Detalle Plantas` sheet name.** Mexico has this (Spanish). Brazil may use a Portuguese equivalent ("Relatório Detalhado das Plantas" or similar). If so, update the fallback in `_pick_sheet()` or rename the sheet lookup.

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
| Add a Brazilian plant                      | `plants.py` — new `PlantSpec` + `REGISTRY` |
| Change which template sheet a line writes  | `plants.py` — `Job.sheet`                  |
| Fix a wrong number for a line              | `plants.py` — `Job.rdp_row` or `pond_line` |
| Translate RDP headers                      | `engine.py` — `MANDATORY_HEADERS`          |
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

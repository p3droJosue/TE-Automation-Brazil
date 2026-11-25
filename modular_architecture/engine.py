"""TE template fill engine — Brazil.

Two halves:

1. **Planning (pure Python, no Excel).** ``load_brazil_db`` reads the four
   Power BI exports in the databases folder; ``plan_plant`` turns them plus a
   ``PlantSpec`` into a flat list of cell writes. Nothing here needs Excel,
   so it can be tested on any machine (see ``analysis/check_against_golden.py``).

2. **Applying (Excel via xlwings).** ``fill_plant`` opens the blank template in
   Excel, applies the writes, recalculates and saves. OneDrive staging and
   the mtime safety net are unchanged from the Mexico engine.

Template layout (identical to Mexico): each product sheet holds up to 15
line blocks, 8 columns wide. Rows below are for the Default layout
(Extruded / Other / Stax / Cookies / Wafer); TC blocks sit 1 row lower and
PC blocks 10 rows lower.
"""
from __future__ import annotations

import calendar
import os
import re
import shutil
import tempfile
import unicodedata
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence, Tuple

# =========================
# Helpers
# =========================
def _norm(s: Any) -> str:
    """Upper-case, strip accents and collapse whitespace."""
    if s is None:
        return ""
    s = "".join(ch for ch in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(ch))
    return " ".join(s.upper().split())


BLOCK_WIDTH = 8

# Column letters inside one block (block 0 starts at column B = 2)
COL = {"B": 2, "C": 3, "D": 4, "E": 5, "F": 6, "G": 7, "H": 8}

# Sheets that hold line blocks. "Stax" is used by Brazil (Sorocaba Canister).
_PRODUCT_SHEETS = ("PC", "TC", "Extruded", "Other", "Cookies or Biscuits", "Wafer", "Stax")


def kind_of(sheet: str) -> str:
    return "PC" if sheet == "PC" else "TC" if sheet == "TC" else "DEFAULT"


ROW_OFFSET = {"DEFAULT": 0, "TC": 1, "PC": 10}

# =========================
# Loss category -> template cell
# =========================
# (column letter, row in DEFAULT layout). The DB exports already use the
# global TE category names, so this table is the whole "translation".
CALENDAR_CELL = ("C", 18)          # Total calendar time hours = days in month x 24

CATEGORY_CELLS: Dict[str, Tuple[str, int]] = {
    # --- left column: time tracking ---
    "Dark hours: line available to run if needed":      ("C", 19),
    "Dark hrs: line available to run if needed":        ("C", 19),   # MES spelling
    "Planned maintenance hours":                        ("C", 22),
    "Planned sanitation hours":                         ("C", 23),
    "Planned startup & shutdown hours":                 ("C", 24),
    "Planned changeovers hours":                        ("C", 25),
    "Planned network outages hours":                    ("C", 26),
    "Team meetings hours":                              ("C", 27),
    "Other scheduled downtime hours":                   ("C", 28),
    "Maintenance overrun hours":                        ("C", 30),
    "Sanitation overrun hours":                         ("C", 31),
    "Changeover overrun hours":                         ("C", 32),
    "Raw material shortages hours":                     ("C", 35),
    "Quality issues hours":                             ("C", 36),
    "Other hours":                                      ("C", 39),
    # --- right column: dark-hour loss tree ---
    "Dark hour Loss Tree - RM shortage":                ("G", 20),
    "Dark hour Loss Tree - FG warehouse":               ("G", 21),
    "Dark hour Loss Tree - Holiday & Religious reasons": ("G", 22),
    "Dark hour Loss Tree - CAPEX & Test (R&D)":         ("G", 23),
    "Dark hour Loss Tree - Staffling":                  ("G", 24),
    "Dark hour Loss Tree - Staffing":                   ("G", 24),
    "Dark hour Loss Tree - Inventory":                  ("G", 27),
    "Dark hour Loss Tree - Other":                      ("G", 28),
    "Dark hour Loss Tree - Corn Cook Hours":            ("G", 29),   # TC only -> G30
    # --- right column: breakdowns & utilities ---
    "Processing Breakdown - Technician Failure":        ("G", 32),
    "Processing Breakdown - Operational Failure":       ("G", 33),
    "Packaging Breakdown - Technician Failure":         ("G", 36),
    "Packaging Breakdown - Operational Failure":        ("G", 37),
    "Utilities - Internal":                             ("G", 40),
    "Utilities - External":                             ("G", 41),
    # --- MES coarse buckets (MES has no sub-split; see BRAZIL_ANALYSIS.md §5) ---
    # Goldens put MES Packaging all on Technical and MES Utilities all on External.
    "Packaging Breakdown":                              ("G", 36),
    "Utilities hours":                                  ("G", 41),
    # Assumptions pending Brazil team confirmation:
    "Dark hrs: unavailable due to constraints":         ("G", 28),   # -> dark loss tree "Other"
    "Processing Breakdown":                             ("G", 32),   # -> Technical failure
    "FG warehouse":                                     ("E", 32),   # unscheduled FG warehouse box
    "Unspecified":                                      ("C", 39),   # -> Other hours
}
# Categories in the exports that the template calculates itself.
IGNORED_CATEGORIES = {"PRODUCTIVE TIME", "TOTAL"}

# Only Corn Cook is TC-specific; on other sheets that row does not exist.
_TC_ONLY = {"Dark hour Loss Tree - Corn Cook Hours"}

_CATEGORY_LOOKUP = {_norm(k): k for k in CATEGORY_CELLS}

# Product table / production table layouts per kind:
#   (first product row, first production row, number of product rows)
PRODUCT_LAYOUT = {
    "DEFAULT": (44, 60, 12),
    "TC":      (45, 69, 20),
    "PC":      (54, 88, 30),
}
# Product-table columns the engine may write. Column G is a formula on PC/TC
# (rated throughput lookups) and on Stax (VLOOKUP on the line type), so it
# is left alone there; column H is always a formula.
PRODUCT_COLS = {"PC": "BCDEF", "TC": "BCDEF", "Stax": "BCDEF"}
DEFAULT_PRODUCT_COLS = "BCDEFG"

# Header row per kind (line name dropdown)
HEADER_ROW = {"DEFAULT": 4, "TC": 4, "PC": 6}
TYPE_CELL = ("C", 7)                  # line type dropdown (Default layout only)
TE_TARGET_ROW = {"DEFAULT": 80, "TC": 97, "PC": 126}   # column G
PC_FRYER_ROWS = (10, 11, 12)          # wash water temp, moisture aim, surface water (column C)


# =========================
# Models
# =========================
@dataclass(frozen=True)
class SKU:
    """One row of a line's product table (master data, not in the DB)."""
    name: str
    ptype: Any = None        # column C — product type
    seasoning: Any = None    # column D — % seasoning
    oil: Any = None          # column E — % oil spray (PC/TC)
    col_f: Any = None        # column F — non-IPS throughput (PC/TC) / extra (Default)
    rated: Any = None        # column G — rated throughput / max kgs
    db_keys: Tuple[Tuple[Optional[str], Optional[str]], ...] = ()  # (Corte, Sabor) pairs feeding this row


@dataclass(frozen=True)
class Job:
    sheet: str                    # template sheet
    cons: int                     # block index (0..14)
    header_text: str              # line name as in the template dropdown
    source: str                   # "APP" | "ITA" | "MES" — which downtime export
    db_line: str                  # 'Linha' in the exports
    type_value: Optional[str] = None      # row 7 line type (Default layout only)
    te_target: Optional[float] = None     # TE % target
    fryer: Optional[Tuple[Any, Any, Any]] = None  # PC: wash water temp, moisture aim, surface water
    skus: Tuple[SKU, ...] = ()

    @property
    def kind(self) -> str:
        return kind_of(self.sheet)


@dataclass(frozen=True)
class PlantSpec:
    key: str
    jobs: List[Job]


@dataclass
class Write:
    sheet: str
    row: int
    col: int
    value: Any


# =========================
# Reading the Brazil exports
# =========================
@dataclass
class BrazilDB:
    # downtime[source][LINE][category] = hours
    downtime: Dict[str, Dict[str, Dict[str, float]]] = field(default_factory=dict)
    # production[LINE] = [(corte, sabor, kg), ...]
    production: Dict[str, List[Tuple[Optional[str], Optional[str], float]]] = field(default_factory=dict)
    files: Dict[str, str] = field(default_factory=dict)   # source -> filename
    filters: Dict[str, str] = field(default_factory=dict)  # source -> "Applied filters" text


_SOURCE_BY_VALUE_LABEL = {"TEMPO_DBL": "APP", "HORAS ITA": "ITA", "DT HORAS": "MES"}


def _read_rows(path: str) -> List[Tuple[Any, ...]]:
    import openpyxl  # local import keeps xlwings-only installs light
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        return [tuple(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()


def _filters_text(rows: Sequence[Tuple[Any, ...]]) -> str:
    for r in rows:
        if r and isinstance(r[0], str) and r[0].startswith("Applied filters"):
            return r[0]
    return ""


def load_brazil_db(folder: str) -> BrazilDB:
    """Read every .xlsx export in *folder* and classify it by content:

    * downtime pivot: A1='Linha', row 2 holds the value label
      (TEMPO_DBL -> APP, HORAS ITA -> ITA, DT Horas -> MES)
    * production list: header row Planta | Linha | Corte | Sabor | Produção Real KG
    """
    if not os.path.isdir(folder):
        raise ValueError(f"Databases folder not found: {folder}")
    db = BrazilDB()
    for fname in sorted(os.listdir(folder)):
        if not fname.lower().endswith((".xlsx", ".xlsm")) or fname.startswith("~$"):
            continue
        rows = _read_rows(os.path.join(folder, fname))
        if len(rows) < 2 or not rows[0]:
            continue
        head = [_norm(v) for v in rows[0]]

        if head[:2] == ["PLANTA", "LINHA"]:
            for r in rows[1:]:
                if len(r) < 5 or not r[1] or not isinstance(r[4], (int, float)):
                    continue
                db.production.setdefault(_norm(r[1]), []).append(
                    (r[2] if r[2] else None, r[3] if r[3] else None, float(r[4])))
            db.files["PROD"] = fname
            db.filters["PROD"] = _filters_text(rows)
            continue

        if head[0] == "LINHA":
            label = next((_norm(v) for v in rows[1][1:] if v), "")
            source = _SOURCE_BY_VALUE_LABEL.get(label)
            if source is None:
                continue
            if source in db.files:
                raise ValueError(f"Two {source} exports in the databases folder: "
                                 f"{db.files[source]} and {fname}")
            lines = [_norm(v) for v in rows[0][1:]]
            data: Dict[str, Dict[str, float]] = {}
            for r in rows[2:]:
                cat = r[0]
                if not isinstance(cat, str) or cat.startswith("Applied filters"):
                    continue
                for ln, v in zip(lines, r[1:]):
                    if ln and ln != "TOTAL" and isinstance(v, (int, float)):
                        data.setdefault(ln, {})[cat.strip()] = float(v)
            db.downtime[source] = data
            db.files[source] = fname
            db.filters[source] = _filters_text(rows)

    missing = [s for s in ("APP", "ITA", "MES", "PROD") if s not in db.files]
    if missing:
        names = {"APP": "Detalhe True Efficiency - APP", "ITA": "Detalhe True Efficiency - ITA",
                 "MES": "Detalle True Efficiency - MES", "PROD": "Detalle True Efficiency (production)"}
        raise ValueError("Databases folder is missing: " + ", ".join(names[m] for m in missing))
    return db


# =========================
# Calendar hours
# =========================
_MONTHS = {m: i for i, m in enumerate(
    ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"], start=1)}


def calendar_hours_from_filename(path: str) -> Optional[float]:
    """'Snk_TE_ITU_P11 2026 OCT(P10 UK).xlsm' -> 31 * 24 = 744."""
    m = re.search(r"(20\d\d)\s*([A-Za-z]{3})", os.path.basename(path))
    if not m or m.group(2).upper() not in _MONTHS:
        return None
    year, month = int(m.group(1)), _MONTHS[m.group(2).upper()]
    return float(calendar.monthrange(year, month)[1] * 24)


# =========================
# Planning
# =========================
def _cell(job: Job, col_letter: str, row: int) -> Tuple[int, int]:
    return row + ROW_OFFSET[job.kind], COL[col_letter] + BLOCK_WIDTH * job.cons


def _sku_key(corte: Any, sabor: Any) -> Tuple[str, str]:
    return _norm(corte), _norm(sabor)


def plan_job(db: BrazilDB, job: Job, calendar_hours: float) -> Tuple[List[Write], List[str]]:
    writes: List[Write] = []
    warns: List[str] = []
    s, kind, o = job.sheet, job.kind, BLOCK_WIDTH * job.cons
    tag = f"{job.sheet} / {job.header_text}"

    def put(row: int, col: int, value: Any) -> None:
        # Only real values are written. Output always starts from a fresh copy
        # of the blank template, so there is nothing stale to clear — and
        # clearing would hit merged cells (e.g. D:E on Other product rows),
        # which Excel refuses.
        if value is not None and value != "":
            writes.append(Write(s, row, col, value))

    # 1) header + static per-line values
    put(HEADER_ROW[kind], COL["B"] + o, job.header_text)
    if kind == "DEFAULT" and job.type_value is not None:
        put(TYPE_CELL[1], COL[TYPE_CELL[0]] + o, job.type_value)
    if job.te_target is not None:
        put(TE_TARGET_ROW[kind], COL["G"] + o, job.te_target)
    if kind == "PC" and job.fryer:
        for r, v in zip(PC_FRYER_ROWS, job.fryer):
            put(r, COL["C"] + o, v)

    # 2) hours
    put(*_cell(job, *CALENDAR_CELL), calendar_hours)
    src = db.downtime.get(job.source, {})
    line_data = src.get(_norm(job.db_line))
    if line_data is None:
        warns.append(f"{tag}: line '{job.db_line}' not found in the {job.source} export — hours left blank.")
        line_data = {}
    values: Dict[Tuple[int, int], float] = {}
    for col_letter, row in set(CATEGORY_CELLS.values()):
        if row == 29 and kind != "TC":
            continue
        values[_cell(job, col_letter, row)] = None  # type: ignore[assignment]
    for cat, hours in line_data.items():
        key = _CATEGORY_LOOKUP.get(_norm(cat))
        if key is None:
            if _norm(cat) not in IGNORED_CATEGORIES:
                warns.append(f"{tag}: unknown category '{cat}' ({hours:.2f} h) not written.")
            continue
        if key in _TC_ONLY and kind != "TC":
            warns.append(f"{tag}: '{cat}' ({hours:.2f} h) only exists on TC — not written.")
            continue
        rc = _cell(job, *CATEGORY_CELLS[key])
        values[rc] = (values.get(rc) or 0.0) + hours
    for (r, c), v in sorted(values.items()):
        put(r, c, v)

    # 3) products + production
    p1, p2, nrows = PRODUCT_LAYOUT[kind]
    skus = list(job.skus)
    by_key: Dict[Tuple[str, str], int] = {}
    for i, sku in enumerate(skus):
        for k in sku.db_keys:
            by_key[_sku_key(*k)] = i
    prod = [0.0] * len(skus)
    extra: List[Tuple[str, float]] = []
    for corte, sabor, kg in db.production.get(_norm(job.db_line), []):
        i = by_key.get(_sku_key(corte, sabor))
        if i is None and len(skus) == 1 and not skus[0].db_keys:
            i = 0  # single-SKU line (e.g. Amacoco: no Corte/Sabor in the export)
        if i is None:
            label = " ".join(str(x) for x in (corte, sabor) if x) or job.db_line
            extra.append((label, kg))
            warns.append(f"{tag}: SKU '{label}' ({kg:,.0f} kg) is not in the SKU master — "
                         f"added as a new row; fill its type/seasoning/throughput by hand.")
        else:
            prod[i] += kg
    if len(skus) + len(extra) > nrows:
        raise ValueError(f"{tag}: {len(skus) + len(extra)} SKUs do not fit the {nrows}-row product table.")
    if job.skus and not db.production.get(_norm(job.db_line)):
        warns.append(f"{tag}: no production rows for line '{job.db_line}' in the production export.")

    for i in range(nrows):
        r1, r2 = p1 + i, p2 + i
        if i < len(skus):
            k = skus[i]
            row_vals = [k.name, k.ptype, k.seasoning, k.oil, k.col_f, k.rated]
            kg = prod[i] or None
        elif i - len(skus) < len(extra):
            name, kg = extra[i - len(skus)]
            row_vals = [name, None, None, None, None, None]
        else:
            row_vals = [None] * 6
            kg = None
        cols = PRODUCT_COLS.get(job.sheet, DEFAULT_PRODUCT_COLS)
        for col_letter, v in zip("BCDEFG", row_vals):
            if col_letter in cols:
                put(r1, COL[col_letter] + o, v)
        put(r2, COL["C"] + o, kg)
    return writes, warns


def plan_plant(db: BrazilDB, plant: PlantSpec, calendar_hours: float) -> Tuple[List[Write], List[str]]:
    writes: List[Write] = []
    warns: List[str] = []
    for job in plant.jobs:
        w, msgs = plan_job(db, job, calendar_hours)
        writes += w
        warns += msgs
    return writes, warns


# =========================
# Engine runner (Excel)
# =========================
def default_out_path(blank_path: str) -> str:
    folder = os.path.dirname(blank_path)
    base = os.path.basename(blank_path)
    lower = base.lower()
    if "_blank" in lower:
        idx = lower.index("_blank")
        base = base[:idx] + "_Filled" + base[idx + 6:]
    else:
        root, ext = os.path.splitext(base)
        base = f"{root}_Filled{ext}"
    return os.path.join(folder, base)


def _is_onedrive_path(path: str) -> bool:
    """OneDrive's sync agent acquires file locks while a workbook is open in
    Excel — when an automated save fires those operations COM-throw
    "Cannot access '<file>'". Detect any OneDrive folder, both personal
    ("OneDrive") and tenant-scoped ("OneDrive - <Org>"), by substring.
    """
    return "onedrive" in (path or "").lower()


def fill_plant(db_path: str, blank_path: str, out_path: str, plant: PlantSpec,
               visible: bool = False) -> List[str]:
    """Fill one plant template. *db_path* is the **databases folder** holding
    the four Brazil exports. Returns a list of warnings (may be empty).

    The workbook is staged through a local temp folder when any path lives
    on OneDrive. Excel COM occasionally throws "Cannot access '<file>'"
    after the workbook has already been written to disk; to avoid flagging
    those runs as failed, we snapshot the output file's mtime before
    starting and treat an exception as success if the file was rewritten.
    """
    # Plan first — pure Python, so data problems fail fast without Excel.
    db = load_brazil_db(db_path)
    hours = calendar_hours_from_filename(blank_path)
    if hours is None:
        raise ValueError("Could not read the month from the template filename "
                         "(expected e.g. '..._P11 2026 OCT(P10 UK).xlsm').")
    writes, warnings = plan_plant(db, plant, hours)

    mtime_before = (
        os.path.getmtime(out_path) if os.path.exists(out_path) else 0.0
    )

    try:
        if _is_onedrive_path(blank_path) or _is_onedrive_path(out_path):
            with tempfile.TemporaryDirectory(prefix="te_filler_") as tmpdir:
                local_tpl = os.path.join(tmpdir, "tpl_" + os.path.basename(blank_path))
                shutil.copy2(blank_path, local_tpl)
                _fill_plant_core(local_tpl, local_tpl, writes, visible)
                shutil.copy2(local_tpl, out_path)
        else:
            _fill_plant_core(blank_path, out_path, writes, visible)
    except Exception:
        # If the output file's mtime advanced during this call the bytes
        # are on disk — whatever COM error fired afterwards (OneDrive lock
        # races, Excel close hiccups, etc.) shouldn't mark the run failed.
        try:
            if (os.path.exists(out_path)
                    and os.path.getmtime(out_path) > mtime_before + 0.5):
                return warnings
        except OSError:
            pass
        raise
    return warnings


def _fill_plant_core(blank_path: str, out_path: str, writes: List[Write], visible: bool = False) -> None:
    import xlwings as xw  # Windows + Excel only

    app = xw.App(visible=visible, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        wb_tpl = app.books.open(blank_path, update_links=False)
        app.calculation = "manual"   # hundreds of writes — recalc once at the end
        sheets = {s.name: s for s in wb_tpl.sheets}
        for w in writes:
            if w.sheet not in sheets:
                raise ValueError(f"Template has no sheet '{w.sheet}'.")
            sheets[w.sheet].range((w.row, w.col)).value = w.value
        app.calculation = "automatic"
        app.api.CalculateFull()
        wb_tpl.save(out_path)

        # File is on disk. Close is best-effort: when the workbook lives on
        # OneDrive the sync agent often grabs a lock the instant save finishes,
        # and Excel's Close() then COM-throws "Cannot access '<file>'". Letting
        # that propagate would mark a correctly-saved plant as Failed.
        try:
            wb_tpl.close()
        except Exception:
            pass

    finally:
        # Each call may COM-throw if Excel is in a modal/busy state from an
        # earlier failure. Swallow individually so app.quit() always runs —
        # otherwise the orphaned Excel process locks files for the next plant.
        for cleanup in (
            lambda: setattr(app, "screen_updating", True),
            lambda: setattr(app, "display_alerts", True),
            app.quit,
        ):
            try:
                cleanup()
            except Exception:
                pass


def apply_writes_openpyxl(blank_path: str, out_path: str, writes: List[Write]) -> None:
    """Developer/test helper: apply a plan with openpyxl (no Excel, no
    recalculation). Use only to compare input cells against a golden — the
    real fill always goes through Excel so formulas and macros stay intact."""
    import openpyxl
    wb = openpyxl.load_workbook(blank_path, keep_vba=True)
    for w in writes:
        wb[w.sheet].cell(w.row, w.col).value = w.value
    wb.save(out_path)

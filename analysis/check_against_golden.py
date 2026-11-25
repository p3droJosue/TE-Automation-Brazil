"""Dry-run the Brazil engine against the golden templates (no Excel needed).

    python analysis/check_against_golden.py                 # all plants, summary
    python analysis/check_against_golden.py SOROCABA -v     # one plant, every mismatch

For each plant it builds the write plan from databases/ and checks it
against the blank and completed templates:

  * SAFETY   — no write may land on a formula cell of the blank template
               (that would destroy the template's calculations) or inside a
               merged area (Excel refuses it). Must be 0.
  * MATCH    — written value equals the golden value
  * DIFF     — written value differs from the golden (expected while the DB
               exports exclude 27-31 Oct; see BRAZIL_ANALYSIS.md)
  * MANUAL   — input cells the analyst filled in the golden that the engine
               does not write (rate losses, gross solids, ...)
"""
from __future__ import annotations

import os
import sys
from collections import defaultdict

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "modular_architecture"))

import engine  # noqa: E402
from plants import REGISTRY, find_templates  # noqa: E402

BLANKS = os.path.join(ROOT, "templates", "TE_blank_templates")
GOLDEN = os.path.join(ROOT, "templates", "TE_completed_templates")


def grid(path, sheets, data_only):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=data_only)
    out = {}
    for s in sheets:
        cells = {}
        for r, row in enumerate(wb[s].iter_rows(max_row=170, max_col=125, values_only=True), start=1):
            for c, v in enumerate(row, start=1):
                if v is not None and v != "":
                    cells[(r, c)] = v
        out[s] = cells
    wb.close()
    return out


def same(a, b):
    if a in (None, "") and b in (None, ""):
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) < 1e-6 * max(1.0, abs(b))
    return engine._norm(a) == engine._norm(b)


def main(argv):
    verbose = "-v" in argv
    only = [a.upper() for a in argv if not a.startswith("-")]
    db = engine.load_brazil_db(os.path.join(ROOT, "databases"))
    total = defaultdict(int)
    for key, blank_path in find_templates(BLANKS):
        if only and key not in only:
            continue
        golden_path = os.path.join(GOLDEN, os.path.basename(blank_path).replace(".xlsm", "_complete.xlsm"))
        plant = REGISTRY[key]
        writes, warns = engine.plan_plant(db, plant, engine.calendar_hours_from_filename(blank_path))
        sheets = sorted({w.sheet for w in writes})
        bf = grid(blank_path, sheets, data_only=False)
        gv = grid(golden_path, sheets, data_only=True)
        gf = grid(golden_path, sheets, data_only=False)

        stats = defaultdict(int)
        lines = []
        # Excel refuses writes into the middle of a merged area
        wb_full = openpyxl.load_workbook(blank_path)
        for w in writes:
            if type(wb_full[w.sheet].cell(w.row, w.col)).__name__ == "MergedCell":
                stats["SAFETY"] += 1
                lines.append(f"  SAFETY  {w.sheet}!{openpyxl.utils.get_column_letter(w.col)}{w.row}: inside a merged area")
        del wb_full
        written = set()
        for w in writes:
            written.add((w.sheet, w.row, w.col))
            b = bf[w.sheet].get((w.row, w.col))
            g = gv[w.sheet].get((w.row, w.col))
            ref = f"{w.sheet}!{openpyxl.utils.get_column_letter(w.col)}{w.row}"
            if isinstance(b, str) and b.startswith("="):
                stats["SAFETY"] += 1
                lines.append(f"  SAFETY  {ref}: blank has formula {b[:40]!r}")
            elif same(w.value, g):
                if w.value not in (None, ""):
                    stats["MATCH"] += 1
            else:
                stats["DIFF"] += 1
                if verbose:
                    lines.append(f"  DIFF    {ref}: engine={w.value!r} golden={g!r}")
        # golden input cells the engine never writes
        for s in sheets:
            for rc, v in gf[s].items():
                if (s,) + rc in written or bf[s].get(rc) == v:
                    continue
                if isinstance(v, str) and v.startswith("="):
                    continue
                if isinstance(v, str) and not v.strip():
                    continue
                stats["MANUAL"] += 1
                if verbose:
                    lines.append(f"  MANUAL  {s}!{openpyxl.utils.get_column_letter(rc[1])}{rc[0]} = {v!r}")
        for k, v in stats.items():
            total[k] += v
        print(f"{key:18} writes={len(writes):4}  match={stats['MATCH']:4}  diff={stats['DIFF']:4}  "
              f"manual={stats['MANUAL']:3}  SAFETY={stats['SAFETY']}  warnings={len(warns)}")
        for m in warns:
            print("  WARN   ", m)
        for ln in lines:
            print(ln)
    print(f"\nTOTAL match={total['MATCH']} diff={total['DIFF']} manual={total['MANUAL']} SAFETY={total['SAFETY']}")
    return 1 if total["SAFETY"] else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))

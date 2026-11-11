# TE-Automation-Brazil

Fills PepsiCo Brazil's True Efficiency (TE) reporting templates from the monthly TE database.
Forked from the Mexico implementation ([TE-Automation](https://github.com/p3droJosue/TE-Automation)).
See `CLAUDE.md` for architecture, onboarding and the Brazil-specific work.

- `modular_architecture/` — app code (`app.py`, `engine.py`, `plants.py`, `build.spec`)
- `databases/` — Brazilian monthly TE databases (`.xlsm`)
- `templates/` — Brazilian blank and completed (golden) templates, one folder per month

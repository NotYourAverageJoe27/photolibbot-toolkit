# PhotoLibBot Toolkit (picture-catalog)

A small toolkit for cataloging, de-duplicating, date-organizing and verifying a large
personal photo/video collection spread across drives, phones and old backups.

- `scripts/pcat.py` – main tool (scan/catalog, duplicate detection incl. video look-alikes,
  date repair, event/trip folder planning, apply/undo, verify, compare new sources).
- `scripts/pcat-win.ps1` – Windows PowerShell 5.1 companion for drives a Linux/Python shell
  cannot reach (inventory, SHA-1 verified copy/move, tree verification).
- `scripts/RUN_template.cmd` – double-clickable launcher template for `pcat-win.ps1`.
- `scripts/selftest.py`, `scripts/_testlib.py` – end-to-end self-test on synthetic files.
- `references/config-example.json` – example config for `pcat.py plan` (all values are placeholders).

## Requirements

- Python 3 with `Pillow`, `numpy`, `openpyxl`
- `ffmpeg` and `ffprobe` on the PATH
- Optional: `pillow-heif` (HEIC photos), `reverse_geocoder` (offline place names)
- Windows PowerShell 5.1 for `pcat-win.ps1` (Windows only)

## First steps

1. Copy the `scripts/` folder to a working folder (not inside your photo library).
2. Check dependencies: `python3 pcat.py doctor`
3. Run the self-test: `python3 selftest.py` — it must print **SELFTEST PASSED** before you
   point the tool at any real files.
4. Copy `references/config-example.json` next to your catalog, replace every placeholder
   (destination folder, homes, side folders, overrides) with your own values.
5. Always start with read-only steps (scan, catalog, dupes, plan) and review the CSV/XLSX
   output before applying anything.

## Safety model

- **Never overwrites**: an existing destination file is never replaced.
- **Quarantine before delete**: duplicates are moved to a quarantine folder, not deleted;
  you delete the quarantine yourself after checking.
- **Logs and undo CSVs**: every apply step writes a log and an undo CSV so the original
  layout can be restored.
- **SHA-1 verify**: copies and moves are verified by SHA-1 hash of source vs. destination;
  a source is only removed after a matching hash.

## Disclaimer

Use at your own risk. This software is provided as is, without warranty of any kind.
Keep independent backups of your photos and videos before running any step that moves files.

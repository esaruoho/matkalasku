# matkalasku

Fill, sign and bake a Finnish travel-expense claim (**Matkalasku**) as a clean,
one-page PDF — from the command line, with no cloud and no account.

You tell it where you drove (there and back) and on which days; it fills the union
template, totals `km × the year's rate`, drops in your signature and today's date, and
bakes the PDF. Your personal details live in a local **`.env` that is gitignored**, so
nothing — bank account, address, car reg, signature — is ever shared or committed.

```
12.6.2026   Helsinki–Kemiö   189 km
15.6.2026   Kemiö–Helsinki   189 km
                              378 km × 0.55 €/km = 207.90 €
```

## Quick start

```bash
git clone https://github.com/esaruoho/matkalasku
cd matkalasku
python3 matkalasku.py          # 1st run: creates .env, asks you to fill it
$EDITOR .env                   # add your name, IBAN, car reg, home city, signature
cp ~/my-signature.png signature.png   # optional — white background is fine
python3 matkalasku.py          # now it asks where/when and bakes the PDF
```

Output lands in `out/` as `yyyy_mm_dd_<handle>_matkalasku_<place>.{xlsx,pdf,json}`.

## What `.env` holds (created + labelled on first run)

| key | meaning |
|-----|---------|
| `NAME` | your full name (the payee) |
| `HANDLE` | short handle for the filename, e.g. `esaruoho` |
| `IBAN` | your bank account number |
| `CAR_REG` | your car registration / plate |
| `HOME_CITY` | where you travel from |
| `SIGNATURE_PATH` | path to your signature PNG (optional) |

**`.env`, `signature.png`, `.routes.json` and everything in `out/` are gitignored.**
Only placeholders (`.env.example`) ship in the repo.

## How it models a trip

Each **travel date is one one-way leg** of the distance you give — drive there one day,
back another → two rows (`A–B`, then `B–A`). For a same-day round trip, give the date
twice. The km rate is stored **per year** in `rates.json` (2025 = 0.59, 2026 = 0.55) and
written into the spreadsheet, so a new year just needs its rate entered once (you're
asked, and it's remembered). Distances are remembered per route, and spelling variants
(Kimiö / Kimito / Kemiö) collapse to one.

## Flags (non-interactive)

```bash
python3 matkalasku.py --to Kemiö --place "Synthcamp 2026" \
  --date 2026-06-12 --date 2026-06-15 --km 189 --yes
# also: --from --purpose --regnr --name --iban --rate --auto-km
#       --perdiem koko:3   (optional per-diem; off by default)
#       --invoice-date --out --no-open
```

## Using a different template

The tool isn't hardwired to one spreadsheet — it reads a **template profile** that says
*where each field goes*. Each `.xlsx` may have a sidecar `<name>.profile.json`; the
bundled one ships at `templates/matkalasku-2026.profile.json`:

```json
{
  "sheet": "Matkalasku 2026",
  "name_cell": "B3", "iban_cell": "B5", "invoice_date_cell": "B38",
  "km": { "first_row": 10, "last_row": 34,
          "cols": {"date":"A","purpose":"B","route":"C","reg":"D","km":"E","type":"F"},
          "type_value": "Kilometrikorvaus" },
  "rate_cell": "H5", "foreign_sheet": "Ulkomaan päivärahat 2026",
  "print_area": "$A$1:$G$40", "print_area_with_perdiem": "$A$1:$G$67",
  "signature": {"col":2,"row":36,"row_off":110000,"cx":1257000,"cy":330000},
  "branding_cells": ["F1","F2"], "remove_branding": true,
  "perdiem": { "first_row": 48, "last_row": 60,
               "cols": {"span":"A","kohde":"C","syy":"E","type":"F"}, "types": {...} }
}
```

To use another form (a 2027 layout, a different union's template, anything):

```bash
python3 matkalasku.py --template templates/my-form.xlsx ...
```

Drop `my-form.xlsx` in `templates/` and write `my-form.profile.json` beside it pointing
at that form's cells — **no code changes.** If a template has no sidecar, the built-in
default (the 2026 layout) is assumed. Note this remaps cells within an `.xlsx`; a wholly
different *file format* (ODS, PDF form) would need its own filler.

**Full field-by-field explanation of the profile format: [`templates/PROFILE.md`](templates/PROFILE.md).**

## Dependencies

- **Core**: Python standard library only.
- **Optional**: [Pillow](https://pypi.org/project/Pillow/) makes a white-background
  signature transparent; **LibreOffice** (`soffice`) bakes the PDF. Without LibreOffice
  you still get the `.xlsx`.

## Reliability notes (please read before sending real claims)

- **Template self-check.** Before filling, the tool reads a few label cells of the
  template and checks they match the profile (`A3`≈"Nimi", `C9`≈"Reitti", …). If a
  template doesn't match its profile it **refuses to fill** rather than put your IBAN in
  the wrong cell. (`--no-verify` bypasses it; don't, unless you know why.)
- **Send the PDF, not the .xlsx.** The € totals are spreadsheet formulas. We strip stale
  caches and set `fullCalcOnLoad`, and LibreOffice recomputes when baking the PDF — so the
  **PDF is authoritative**. An `.xlsx` opened in an app stuck in manual-calc mode could
  show a stale number.
- **Check the rate.** `rates.json` ships `2025=0.59, 2026=0.55`. These are the values used;
  **verify the official kilometrikorvaus for your year at vero.fi** and update `rates.json`
  (or pass `--rate`). The tool trusts what's in the file.
- **`--auto-km` is a hint, not gospel.** It uses public OpenStreetMap/OSRM and routes to a
  place's *centroid*, so it's usually short of your real route. A distance you state always
  wins and is remembered; the auto guess never overwrites it.
- **Tested + CI.** `python -m unittest test_matkalasku` (14 tests); GitHub Actions runs them
  on every push. PDF baking (LibreOffice) is environment-specific and not covered by CI.

## Architecture

`matkalasku_core.py` is the engine (profile + fill + sign + rates/routes); `matkalasku.py`
is the thin `.env` front-end; `xlsx_fill.py` is a stdlib surgical `.xlsx` editor. The same
`matkalasku_core.py` + `xlsx_fill.py` also power [convey](https://github.com/esaruoho/convey)'s
`convey matkalasku` (which uses a vault profile instead of `.env`) — they're kept
byte-identical, so a fix in one engine reaches both.

## Privacy

This tool sends nothing anywhere, except an **optional** anonymous distance lookup
(`--auto-km`) to public OpenStreetMap/OSRM endpoints — which only ever sees place names
you type, never your personal details. Everything else is local files.

---

Extracted from [convey](https://github.com/esaruoho/convey)'s `convey matkalasku`.

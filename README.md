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

## Dependencies

- **Core**: Python standard library only.
- **Optional**: [Pillow](https://pypi.org/project/Pillow/) makes a white-background
  signature transparent; **LibreOffice** (`soffice`) bakes the PDF. Without LibreOffice
  you still get the `.xlsx`.

## Privacy

This tool sends nothing anywhere, except an **optional** anonymous distance lookup
(`--auto-km`) to public OpenStreetMap/OSRM endpoints — which only ever sees place names
you type, never your personal details. Everything else is local files.

---

Extracted from [convey](https://github.com/esaruoho/convey)'s `convey matkalasku`.

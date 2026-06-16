# Template profiles — how a spreadsheet's layout is described

`matkalasku` is **not hardwired** to one spreadsheet. For each template `.xlsx` it reads
a **profile** that says *where every field goes*. That way a different form (a 2027
layout, another union's template) is supported by writing a small JSON file — **no code
changes.**

## Where the profile comes from

For a template `templates/foo.xlsx`, the tool looks for a sidecar next to it:

```
templates/foo.xlsx
templates/foo.profile.json   ← its profile
```

If there's no sidecar, the **built-in default** (the bundled 2026 layout) is used — so
the shipped template works out of the box. The bundled profile lives at
`templates/matkalasku-2026.profile.json` and is the worked example below.

Use another template with:

```bash
python3 matkalasku.py --template templates/foo.xlsx ...
```

## The mental model

Every entry is **"put this piece of data → in this cell."** The left side is *what*
(your name, your IBAN, a kilometre row); the right side is *where* in the grid.

- **Cells** are spreadsheet refs like `B3`, `H5` — column letter + row number.
- **Columns** in a table are given as letters (`"date": "A"`), rows as numbers.
- **Booleans** at the top are *behaviour*, not location (hide a sheet, strip branding…).

To adapt to a new form: open it in a spreadsheet, see which cell holds the name, which
rows are the kilometre table, etc., and copy those refs into a new `*.profile.json`.

## Every field explained

```jsonc
{
  // ── which sheets (tabs) ──
  "sheet": "Matkalasku 2026",                  // the tab we fill
  "foreign_sheet": "Ulkomaan päivärahat 2026", // 2nd tab: holds the €/km rate + country list
  "hide_foreign_sheet": true,                  // hide that 2nd tab → it won't print
  "remove_branding": true,                     // strip the template-maker's logo + print footer
  "fit_to_page": true,                         // scale so it prints on ONE page

  // ── header fields (single cells) ──
  "name_cell": "B3",            // your name → B3
  "iban_cell": "B5",            // your IBAN → B5
  "invoice_date_cell": "B38",   // signing date "Päiväys" → B38
  "name_clarify_cells": ["C40","C67","C95"],  // "Nimen selvennys" name (3 signature blocks)
  "branding_cells": ["F1","F2"],              // cells whose maker-branding text we blank out

  // ── the kilometre rate ──
  "rate_cell": "H5",            // the year's €/km is written here (on foreign_sheet);
                                // every row's € recomputes from it

  // ── what area prints ──
  "print_area": "$A$1:$G$40",              // km claim only (no per-diem)
  "print_area_with_perdiem": "$A$1:$G$67", // larger area used when per-diem rows are added

  // ── the kilometre table (one row per travel day) ──
  "km": {
    "first_row": 10, "last_row": 34,   // the table's row range (here, 25 rows)
    "cols": {                          // which column each field writes into:
      "date": "A",      //  Päiväys
      "purpose": "B",   //  Matkan tarkoitus
      "route": "C",     //  Reitti, e.g. "Vuosaari–Kimiö"
      "reg": "D",       //  Rek.nro (car registration)
      "km": "E",        //  Km-määrä yhteensä
      "type": "F"       //  Km-korvaus (a dropdown)
    },
    "type_value": "Kilometrikorvaus"   // the dropdown value placed in column F
  },

  // ── optional per-diem table (only used with --perdiem) ──
  "perdiem": {
    "first_row": 48, "last_row": 60,
    "cols": { "span": "A",     // date range (Matkan alku ja loppu)
              "kohde": "C",    // destination (Matkakohde)
              "syy": "E",      // reason (Matkan syy)
              "type": "F" },   // per-diem type (drives the € formula)
    "types": {                 // your shorthand on the CLI → the exact dropdown text:
      "koko":   "Kokopäiväraha",
      "osa":    "Osapäiväraha",
      "koko-2": "Kokopäiväraha (vähennetty 2 ilmaista ateriaa)",
      "osa-1":  "Osapäiväraha (vähennetty 1 ilmainen ateria)",
      "ateria": "Ateriakorvaus, kotimaa"
    }
  },

  // ── where the signature image is anchored ──
  "signature": {
    "col": 2,            // column C  (0-based: A=0, B=1, C=2)
    "row": 36,           // row 37    (0-based) — the "Allekirjoitus" label row
    "row_off": 110000,   // nudge down ~half a row so it sits on the line
    "cx": 1257000,       // width  (EMU; 914400 EMU = 1 inch → ~1.4")
    "cy": 330000         // height (~0.36") — small enough to clear "Nimen selvennys"
  }
}
```

### A note on units

- **Cells / rows / columns** are normal spreadsheet refs (1-based), e.g. `B3`, row `10`,
  column `"A"`.
- **The signature** is a floating image, so it's anchored differently: `col`/`row` are
  **0-based** indexes, and `cx`/`cy`/`row_off` are in **EMU** (English Metric Units,
  Excel's drawing unit — **914,400 EMU = 1 inch**). Tweak `row`/`row_off` to move it up
  or down, `cx`/`cy` to resize. After changing a signature anchor, render the PDF and
  eyeball it — that's the only reliable check.

### Optional / behaviour fields

`hide_foreign_sheet`, `remove_branding`, `fit_to_page` are on/off switches. The per-diem
`types` map and `km.type_value` are the exact dropdown strings your form expects (these
must match the form's validation list, character-for-character).

## Limits

This remaps cells **within an `.xlsx`**. A wholly different *file format* (OpenDocument
`.ods`, a fillable PDF form) would need its own filler — `xlsx_fill.py` speaks `.xlsx`
only.

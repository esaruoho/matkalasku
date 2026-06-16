"""Tests for matkalasku_core + xlsx_fill — the standalone Matkalasku engine.

Deterministic, no network, no model tokens. PDF baking (LibreOffice) is NOT exercised
here; we test the .xlsx fill, the maths, the template self-check, and the surgical editor.
"""
import base64
import json
import tempfile
import unittest
import zipfile
from pathlib import Path

import matkalasku_core as core
import xlsx_fill

TEMPLATE = Path(__file__).resolve().parent / "templates" / "matkalasku-2026.xlsx"
PROFILE = core.DEFAULT_PROFILE

_TINY_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


class Helpers(unittest.TestCase):
    def test_normalize_date(self):
        self.assertEqual(core.normalize_date("2026-06-12"), "12.6.2026")
        self.assertEqual(core.normalize_date("12.06.2026"), "12.6.2026")

    def test_fi_date_is_portable(self):
        import datetime
        self.assertEqual(core.fi_date(datetime.date(2026, 6, 5)), "5.6.2026")

    def test_route_leg_and_format_iban(self):
        self.assertEqual(core.route_string("Vuosaari", "Kemiö", False), "Vuosaari–Kemiö")
        self.assertEqual(core.format_iban("FI4813883500076322"), "FI48 1388 3500 0763 22")

    def test_place_aliases(self):
        for s in ("Kimiö", "Kimito", "Kemiönsaari", "kemiö"):
            self.assertEqual(core.canonical_place(s), "Kemiö")

    def test_col_to_index(self):
        self.assertEqual(xlsx_fill.col_to_index("F"), 6)


class CellReader(unittest.TestCase):
    """Regression: self-closing cells must not swallow the next real cell."""

    def test_read_cell_text_handles_self_closing_cells(self):
        got = xlsx_fill.read_cell_text(TEMPLATE, "Matkalasku 2026",
                                       ["A3", "A5", "A9", "C9", "E9"])
        self.assertIn("Nimi", got["A3"])
        self.assertIn("Tilinumero", got["A5"])
        self.assertIn("Päiväys", got["A9"])
        self.assertEqual(got["C9"], "Reitti")


class Model(unittest.TestCase):
    def test_totals(self):
        m = core.Matkalasku(km_per_leg=189, dates=["12.6.2026", "15.6.2026"], rate=0.55)
        self.assertEqual(m.total_km, 378)
        self.assertEqual(m.total_eur, 207.9)

    def test_legs_alternate_and_demo_cleared(self):
        m = core.Matkalasku(name="Esa", iban="FI00", regnr="NJE-521", purpose="Keikka",
                            origin="Vuosaari", destination="Kimiö", km_per_leg=189,
                            dates=["12.6.2026", "15.6.2026"])
        c = m.cells(PROFILE)
        self.assertEqual(c["A10"], "12.6.2026")
        self.assertEqual(c["C10"], "Vuosaari–Kimiö")
        self.assertEqual(c["C11"], "Kimiö–Vuosaari")
        self.assertEqual(c["E10"], 189)
        self.assertIsNone(c["E12"])
        self.assertIsNone(c["F1"]); self.assertIsNone(c["F2"])  # branding cleared

    def test_perdiem_off_by_default_and_fills(self):
        self.assertEqual(core.Matkalasku().perdiem, [])
        m = core.Matkalasku(destination="Kimiö", purpose="K", km_per_leg=189,
                            dates=["12.6.2026", "15.6.2026"], perdiem=["Kokopäiväraha"])
        c = m.cells(PROFILE)
        self.assertEqual(c["F48"], "Kokopäiväraha")
        self.assertEqual(c["A48"], "12.6.2026–15.6.2026")


class RatesAndRoutes(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.rates = Path(self.tmp.name) / "rates.json"
        self.routes = Path(self.tmp.name) / "routes.json"

    def tearDown(self):
        self.tmp.cleanup()

    def test_rates_seed_and_save(self):
        self.assertEqual(core.rate_for_year(self.rates, "2026"), 0.55)
        core.save_rate(self.rates, "2027", 0.57)
        self.assertEqual(core.rate_for_year(self.rates, "2027"), 0.57)

    def test_route_alias_and_no_auto_downgrade(self):
        core.remember_route(self.routes, "Vuosaari", "Kemiö", 189, "stated")
        self.assertEqual(core.lookup_cached_km(self.routes, "Vuosaari", "Kimiö"), 189)  # alias
        core.remember_route(self.routes, "Vuosaari", "Kimito", 158, "auto (osrm)")
        self.assertEqual(core.lookup_cached_km(self.routes, "Vuosaari", "Kemiö"), 189)  # not downgraded


class FillAndVerify(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.out = Path(self.tmp.name) / "out.xlsx"

    def tearDown(self):
        self.tmp.cleanup()

    def test_fill_preserves_parts_and_values_and_recalc_setup(self):
        m = core.Matkalasku(name="Esa Ruoho", iban="FI21 1234 5600 0007 85",
                            regnr="NJE-521", origin="Vuosaari", destination="Kimiö",
                            km_per_leg=189, dates=["12.6.2026", "15.6.2026"], rate=0.55)
        core.fill(m, self.out, PROFILE, TEMPLATE)
        with zipfile.ZipFile(TEMPLATE) as z:
            before = set(z.namelist())
        with zipfile.ZipFile(self.out) as z:
            after = set(z.namelist())
            sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
            wb = z.read("xl/workbook.xml").decode("utf-8")
        self.assertEqual(before, after)           # comments/dropdowns/media preserved
        self.assertIn("Esa Ruoho", sheet)
        self.assertIn("Vuosaari–Kimiö", sheet)
        self.assertIn("_xlfn.IFS", sheet)         # formulas untouched
        self.assertIn("fullCalcOnLoad", wb)       # recompute on open
        self.assertRegex(wb, r'Ulkomaan päivärahat 2026"[^>]*state="hidden"')

    def test_self_check_refuses_a_mismatched_profile(self):
        bad = json.loads(json.dumps(PROFILE))
        bad["verify"] = {"A3": "THIS-LABEL-IS-NOT-IN-THE-TEMPLATE"}
        m = core.Matkalasku(name="X", destination="Kimiö", km_per_leg=10, dates=["1.1.2026"])
        with self.assertRaises(ValueError):
            core.fill(m, self.out, bad, TEMPLATE)
        # verify=False bypasses the guard
        core.fill(m, self.out, bad, TEMPLATE, verify=False)
        self.assertTrue(self.out.exists())

    def test_signature_embedded_when_present(self):
        sig = Path(self.tmp.name) / "sig.png"
        sig.write_bytes(_TINY_PNG)
        m = core.Matkalasku(name="X", destination="Kimiö", km_per_leg=189,
                            dates=["16.6.2026"], invoice_date="16.6.2026")
        core.fill(m, self.out, PROFILE, TEMPLATE, signature=sig)
        with zipfile.ZipFile(self.out) as z:
            sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
            drawing = z.read("xl/drawings/drawing1.xml").decode("utf-8")
            img = z.read("xl/media/image1.png")
        self.assertIn("<drawing ", sheet)
        self.assertIn("signature", drawing)
        self.assertNotIn("hlinkClick", drawing)
        with zipfile.ZipFile(TEMPLATE) as z:
            self.assertNotEqual(img, z.read("xl/media/image1.png"))


if __name__ == "__main__":
    unittest.main()

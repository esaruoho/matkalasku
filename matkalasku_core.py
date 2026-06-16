"""matkalasku_core — the config-agnostic engine for filling a Finnish Matkalasku.

This is the SINGLE SOURCE OF TRUTH for the spreadsheet logic. It knows nothing about
where your identity comes from (a `.env`, a vault profile, flags) or about any CLI —
it just takes data + a template profile and fills/signs/bakes. The front-ends
(matkalasku.py here; convey's `convey matkalasku`) supply config and call into this.

Both repos carry a BYTE-IDENTICAL copy of this file and of xlsx_fill.py; a sync test
fails the instant they diverge. Edit the canonical copy (this repo) and copy it over.

stdlib only for the core. Optional: Pillow (signature transparency + aspect),
LibreOffice `soffice` (PDF).
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path

import xlsx_fill

RATE_FALLBACK = 0.55
DEFAULT_RATES = {"2025": 0.59, "2026": 0.55}

EMU_PER_POINT = 12700  # 1 pt = 12700 EMU; 914400 EMU = 1 inch

# ── template profile: WHERE each field goes (see templates/PROFILE.md) ──
DEFAULT_PROFILE = {
    "sheet": "Matkalasku 2026",
    "foreign_sheet": "Ulkomaan päivärahat 2026",
    "hide_foreign_sheet": True,
    "remove_branding": True,
    "fit_to_page": True,
    "name_cell": "B3",
    "iban_cell": "B5",
    "invoice_date_cell": "B38",
    "name_clarify_cells": ["C40", "C67", "C95"],
    "branding_cells": ["F1", "F2"],
    "rate_cell": "H5",
    "print_area": "$A$1:$G$40",
    "print_area_with_perdiem": "$A$1:$G$67",
    # self-check: these cells must contain these label substrings, or we refuse to fill
    "verify": {"A3": "Nimi", "A5": "Tilinumero", "A9": "Päiväys", "C9": "Reitti", "E9": "Km"},
    "km": {
        "first_row": 10, "last_row": 34,
        "cols": {"date": "A", "purpose": "B", "route": "C", "reg": "D", "km": "E", "type": "F"},
        "type_value": "Kilometrikorvaus",
    },
    "perdiem": {
        "first_row": 48, "last_row": 60,
        "cols": {"span": "A", "kohde": "C", "syy": "E", "type": "F"},
        "types": {
            "koko": "Kokopäiväraha",
            "osa": "Osapäiväraha",
            "koko-2": "Kokopäiväraha (vähennetty 2 ilmaista ateriaa)",
            "osa-1": "Osapäiväraha (vähennetty 1 ilmainen ateria)",
            "ateria": "Ateriakorvaus, kotimaa",
        },
    },
    # signature: anchored at 0-based (col,row); cy auto-fits to fit_to_rows (1-based) if
    # the template's row heights are readable, else the cx/cy below are used as-is.
    "signature": {"col": 2, "row": 36, "row_off": 110000, "fit_to_rows": [38, 39],
                  "cx": 1257000, "cy": 330000},
}


def load_template_profile(template: Path) -> dict:
    """The profile for a template: its `<name>.profile.json` sidecar, or the default."""
    sidecar = Path(template).with_name(Path(template).stem + ".profile.json")
    if sidecar.exists():
        try:
            return json.loads(sidecar.read_text(encoding="utf-8"))
        except Exception as e:  # noqa: BLE001
            raise SystemExit(f"matkalasku: bad template profile {sidecar}: {e}")
    return DEFAULT_PROFILE


# ── place aliases + geocoding ──
PLACE_ALIASES = {
    "kimiö": "Kemiö", "kimio": "Kemiö", "kimito": "Kemiö", "kimitoön": "Kemiö",
    "kimitoon": "Kemiö", "kemiönsaari": "Kemiö", "kemionsaari": "Kemiö",
    "kemiö": "Kemiö", "björkboda": "Kemiö", "bjorkboda": "Kemiö",
}


def canonical_place(name: str) -> str:
    return PLACE_ALIASES.get((name or "").strip().lower(), (name or "").strip())


def geocode_candidates(name: str) -> list:
    out, seen = [], set()
    for n in (name, canonical_place(name)):
        n = (n or "").strip()
        if n and n.lower() not in seen:
            seen.add(n.lower())
            out.append(n)
    return out


def lookup_oneway_km(origin, destination, timeout=8.0):
    """Best-effort one-way driving km via public Nominatim + OSRM. Network, not magic.

    NOTE: OSRM routes to a place's centroid, so this is usually SHORT of your real route
    (Helsinki→Kemiö measured 189 km; OSRM says ~158). Treat it as a hint — a distance
    you state is authoritative and is never overwritten by this guess."""
    import urllib.parse
    import urllib.request

    def geocode(place):
        for cand in geocode_candidates(place):
            q = cand if "," in cand else f"{cand}, Finland"
            url = "https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode(
                {"q": q, "format": "json", "limit": 1})
            req = urllib.request.Request(url, headers={"User-Agent": "matkalasku/1.0"})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                data = json.loads(r.read().decode("utf-8"))
            if data:
                return float(data[0]["lon"]), float(data[0]["lat"])
        return None

    try:
        a, b = geocode(origin), geocode(destination)
        if not a or not b:
            return None
        coords = f"{a[0]},{a[1]};{b[0]},{b[1]}"
        url = f"https://router.project-osrm.org/route/v1/driving/{coords}?overview=false"
        req = urllib.request.Request(url, headers={"User-Agent": "matkalasku/1.0"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            data = json.loads(r.read().decode("utf-8"))
        routes = data.get("routes") or []
        return round(routes[0]["distance"] / 1000.0) if routes else None
    except Exception:
        return None


# ── rates (per year, updatable) — path is injected by the front-end ──
def load_rates(path: Path) -> dict:
    try:
        rates = json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        rates = {}
    if not rates:
        rates = dict(DEFAULT_RATES)
        try:
            Path(path).write_text(json.dumps(rates, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass
    return rates


def rate_for_year(path, year):
    v = load_rates(path).get(str(year))
    return float(v) if v is not None else None


def save_rate(path, year, rate):
    rates = load_rates(path)
    rates[str(year)] = round(float(rate), 4)
    Path(path).write_text(json.dumps(rates, ensure_ascii=False, indent=2), encoding="utf-8")


def latest_known_rate(path):
    rates = load_rates(path)
    if not rates:
        return None
    newest = max(rates, key=str)
    return float(rates[newest]), newest


# ── routes (remembered distances) — path injected ──
def _route_key(o, d):
    return f"{canonical_place(o).lower()}|{canonical_place(d).lower()}"


def load_routes(path) -> dict:
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        return {}


def lookup_cached_km(path, o, d):
    routes = load_routes(path)
    for key in (_route_key(o, d), _route_key(d, o)):
        if key in routes:
            try:
                return float(routes[key]["oneway_km"])
            except Exception:
                continue
    return None


def remember_route(path, o, d, oneway_km, source="stated"):
    routes = load_routes(path)
    key = _route_key(o, d)
    existing = routes.get(key)
    if existing and source.startswith("auto") and not str(existing.get("source", "")).startswith("auto"):
        return
    routes[key] = {"origin": canonical_place(o), "destination": canonical_place(d),
                   "oneway_km": round(float(oneway_km), 1), "source": source}
    Path(path).write_text(json.dumps(routes, ensure_ascii=False, indent=2), encoding="utf-8")


# ── small helpers ──
def normalize_date(s: str) -> str:
    s = s.strip()
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", s)
    if m:
        y, mo, d = m.groups()
        return f"{int(d)}.{int(mo)}.{y}"
    m = re.match(r"^(\d{1,2})\.(\d{1,2})\.(\d{4})$", s)
    if m:
        d, mo, y = m.groups()
        return f"{int(d)}.{int(mo)}.{y}"
    return s


def fi_date(d) -> str:
    """Finnish d.m.yyyy — portable (no platform-specific %-d, which breaks on Windows)."""
    return f"{d.day}.{d.month}.{d.year}"


def route_string(origin, destination, round_trip=True):
    o, d = origin.strip(), destination.strip()
    return f"{o}–{d}–{o}" if round_trip else f"{o}–{d}"


def format_iban(iban: str) -> str:
    s = re.sub(r"\s+", "", iban or "")
    return " ".join(s[i:i + 4] for i in range(0, len(s), 4)) if s else ""


def prepare_signature(path):
    """Return a transparent-background PNG path (white→alpha, cached) or the raw/None."""
    path = Path(path) if path else None
    if not path or not path.exists():
        return None
    out = path.with_name(path.stem + "-transparent.png")
    try:
        if not out.exists() or out.stat().st_mtime < path.stat().st_mtime:
            from PIL import Image
            im = Image.open(path).convert("RGBA")
            im.putdata([(r, g, b, 0) if (r > 235 and g > 235 and b > 235) else (r, g, b, a)
                        for (r, g, b, a) in im.getdata()])
            im.save(out)
        return out
    except Exception:
        return path


def _img_aspect(path):
    try:
        from PIL import Image
        with Image.open(path) as im:
            w, h = im.size
        return w / h if h else None
    except Exception:
        return None


def find_soffice():
    import shutil
    for cand in ("soffice", "libreoffice",
                 "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        hit = shutil.which(cand) if "/" not in cand else (cand if Path(cand).exists() else None)
        if hit:
            return hit
    return None


def to_pdf(xlsx_path, out_dir):
    import subprocess
    soffice = find_soffice()
    if not soffice:
        return None
    try:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir",
                        str(out_dir), str(xlsx_path)], check=True, capture_output=True, timeout=120)
    except Exception:
        return None
    pdf = Path(out_dir) / (Path(xlsx_path).stem + ".pdf")
    return pdf if pdf.exists() else None


@dataclass
class Matkalasku:
    name: str = ""
    iban: str = ""
    purpose: str = ""
    regnr: str = ""
    origin: str = ""
    destination: str = ""
    km_per_leg: float = 0.0
    dates: list = field(default_factory=list)
    rate: float = RATE_FALLBACK
    perdiem: list = field(default_factory=list)
    invoice_date: str = ""

    @property
    def total_km(self):
        return self.km_per_leg * len(self.dates)

    @property
    def total_eur(self):
        return round(self.total_km * self.rate, 2)

    @property
    def route(self):
        return f"{self.origin.strip()}↔{self.destination.strip()}"

    def leg_route(self, i):
        return (route_string(self.origin, self.destination, False) if i % 2 == 0
                else route_string(self.destination, self.origin, False))

    @property
    def date_span(self):
        if not self.dates:
            return ""
        return self.dates[0] if len(self.dates) == 1 else f"{self.dates[0]}–{self.dates[-1]}"

    def cells(self, p: dict):
        km = p["km"]
        kc = km["cols"]
        first, last = km["first_row"], km["last_row"]
        max_trips = last - first + 1
        if len(self.dates) > max_trips:
            raise ValueError(f"{len(self.dates)} dates exceeds the template's {max_trips} rows")
        pd = p.get("perdiem") or {}
        pc = pd.get("cols", {})
        pfirst, plast = pd.get("first_row"), pd.get("last_row")
        max_pd = (plast - pfirst + 1) if pfirst else 0
        if len(self.perdiem) > max_pd:
            raise ValueError(f"{len(self.perdiem)} per-diem rows exceeds {max_pd}")

        cells = {p["name_cell"]: self.name, p["iban_cell"]: self.iban}
        if self.invoice_date and p.get("invoice_date_cell"):
            cells[p["invoice_date_cell"]] = self.invoice_date
        if self.name:
            for c in p.get("name_clarify_cells", []):
                cells[c] = self.name
        for c in p.get("branding_cells", []):
            cells[c] = None
        for r in range(first, last + 1):
            for col in kc.values():
                cells[f"{col}{r}"] = None
        for i, d in enumerate(self.dates):
            r = first + i
            cells[f"{kc['date']}{r}"] = d
            cells[f"{kc['purpose']}{r}"] = self.purpose
            cells[f"{kc['route']}{r}"] = self.leg_route(i)
            cells[f"{kc['reg']}{r}"] = self.regnr
            cells[f"{kc['km']}{r}"] = self.km_per_leg
            cells[f"{kc['type']}{r}"] = km["type_value"]
        if pfirst:
            for r in range(pfirst, plast + 1):
                for col in pc.values():
                    cells[f"{col}{r}"] = None
            for i, laatu in enumerate(self.perdiem):
                r = pfirst + i
                cells[f"{pc['span']}{r}"] = self.date_span
                cells[f"{pc['kohde']}{r}"] = self.destination
                cells[f"{pc['syy']}{r}"] = self.purpose
                cells[f"{pc['type']}{r}"] = laatu
        return cells


def verify_template(template, profile) -> list:
    """Return a list of mismatches between the template and its profile (#2: never fill
    blind). Empty list = the template's label cells are where the profile says."""
    checks = profile.get("verify") or {}
    if not checks:
        return []
    got = xlsx_fill.read_cell_text(template, profile["sheet"], list(checks.keys()))
    problems = []
    for cell, expect in checks.items():
        actual = got.get(cell, "") or ""
        if expect.casefold() not in actual.casefold():
            problems.append(f"{cell}: expected to contain {expect!r}, found {actual!r}")
    return problems


def _signature_tuple(template, profile, sig_file):
    s = profile.get("signature") or {}
    cx, cy = s.get("cx"), s.get("cy")
    rows = s.get("fit_to_rows")
    if rows:
        try:
            heights = xlsx_fill.read_row_heights(template, profile["sheet"], rows)
            pts = sum(heights.get(int(r), 15.0) for r in rows) * 0.92
            cy = int(pts * EMU_PER_POINT)
            aspect = _img_aspect(sig_file) or ((s["cx"] / s["cy"]) if s.get("cx") and s.get("cy") else 3.8)
            cx = int(cy * aspect)
        except Exception:
            cx, cy = s.get("cx"), s.get("cy")
    return (str(sig_file), s["col"], s["row"], cx, cy, s.get("row_off", 0))


def fill(data: Matkalasku, out_path, profile: dict, template, signature=None,
         verify: bool = True) -> Path:
    p = profile
    if verify:
        problems = verify_template(template, profile)
        if problems:
            raise ValueError(
                "template does not match its profile (refusing to fill blind):\n  - "
                + "\n  - ".join(problems)
                + "\nFix the profile's cell refs (see templates/PROFILE.md) or pass verify=False.")
    area = p.get("print_area_with_perdiem", p.get("print_area")) if data.perdiem else p.get("print_area")
    sig = None
    sig_file = prepare_signature(signature) if signature else None
    if sig_file and p.get("signature"):
        sig = _signature_tuple(template, profile, sig_file)
    extra = {}
    if p.get("rate_cell") and p.get("foreign_sheet"):
        extra = {p["foreign_sheet"]: {p["rate_cell"]: data.rate}}
    hide = ([p["foreign_sheet"]] if p.get("hide_foreign_sheet") and p.get("foreign_sheet") else None)
    return xlsx_fill.set_cells(
        template, out_path, p["sheet"], data.cells(p),
        hide_sheets=hide,
        print_area=(p["sheet"], area) if area else None,
        fit_to_page=p.get("fit_to_page", True),
        strip_footer=p.get("remove_branding", True),
        remove_drawing=p.get("remove_branding", True),
        signature=sig,
        extra_cells=extra,
    )

#!/usr/bin/env python3
"""matkalasku — fill, sign and bake a Finnish travel-expense claim (Matkalasku) as a PDF.

Zero magic: you tell it where you drove (there and back) and on which days; it fills the
union template, totals km × the year's rate, drops in your signature and today's date,
and bakes a clean one-page PDF. No cloud, no account — your details live in a local
`.env` that is gitignored, so nothing personal is ever shared.

First run creates a labelled `.env` for you to fill in (name, IBAN, car reg, home city,
signature). Then: `python matkalasku.py` (interactive) or with flags (see --help).

Deps: standard library only for the core. Optional: Pillow (makes a white-background
signature transparent) and LibreOffice (`soffice`, bakes the PDF — without it you get
the .xlsx and recompute it yourself).
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import xlsx_fill

REPO = Path(__file__).resolve().parent
ENV = REPO / ".env"
RATES = REPO / "rates.json"
ROUTES = REPO / ".routes.json"
TEMPLATE = REPO / "templates" / "matkalasku-2026.xlsx"
OUT_DIR = REPO / "out"

# ── the spreadsheet's shape (the bundled 2026 union template) ──
SHEET = "Matkalasku 2026"
FOREIGN_SHEET = "Ulkomaan päivärahat 2026"
PRINT_AREA_KM = "$A$1:$G$40"
PRINT_AREA_WITH_PERDIEM = "$A$1:$G$67"
RATE_CELL = "H5"                 # the km-rate cell every row formula references
RATE_FALLBACK = 0.55
DEFAULT_RATES = {"2025": 0.59, "2026": 0.55}
KM_TYPE = "Kilometrikorvaus"
FIRST_ROW, LAST_ROW = 10, 34
MAX_TRIPS = LAST_ROW - FIRST_ROW + 1
PERDIEM_FIRST_ROW, PERDIEM_LAST_ROW = 48, 60
MAX_PERDIEM = PERDIEM_LAST_ROW - PERDIEM_FIRST_ROW + 1
PERDIEM_TYPES = {
    "koko": "Kokopäiväraha",
    "osa": "Osapäiväraha",
    "koko-2": "Kokopäiväraha (vähennetty 2 ilmaista ateriaa)",
    "osa-1": "Osapäiväraha (vähennetty 1 ilmainen ateria)",
    "ateria": "Ateriakorvaus, kotimaa",
}
CELL_NAME, CELL_IBAN = "B3", "B5"
CELL_INVOICE_DATE = "B38"
CELL_NAME_CLARIFY = ("C40", "C67", "C95")
BRANDING_CELLS = ("F1", "F2")
SIG_ANCHOR_COL, SIG_ANCHOR_ROW, SIG_ROW_OFF = 2, 36, 110000
SIG_CY, SIG_CX = 330000, 1257000

PLACE_ALIASES = {
    "kimiö": "Kemiö", "kimio": "Kemiö", "kimito": "Kemiö", "kimitoön": "Kemiö",
    "kimitoon": "Kemiö", "kemiönsaari": "Kemiö", "kemionsaari": "Kemiö",
    "kemiö": "Kemiö", "björkboda": "Kemiö", "bjorkboda": "Kemiö",
}

ENV_TEMPLATE = """\
# ─────────────────────────────────────────────────────────────────────────────
# matkalasku — your details. This file is PRIVATE and gitignored: it is never
# shared or committed. Fill in the values after each `=` (no quotes needed).
# ─────────────────────────────────────────────────────────────────────────────

# Your full name, as the payee on the claim
NAME=

# A short handle used in the output filename (e.g. esaruoho)
HANDLE=

# Your bank account number in IBAN format (e.g. FI00 0000 0000 0000 00)
IBAN=

# Your car registration / license plate (e.g. ABC-123)
CAR_REG=

# Where you travel FROM — your home city or district (e.g. Helsinki)
HOME_CITY=Helsinki

# Path to your signature image (PNG). White background is fine — it is made
# transparent automatically if Pillow is installed. Leave blank for no signature.
SIGNATURE_PATH=signature.png
"""


# ── .env: first-run creation + loading ──

def ensure_env() -> bool:
    """If .env is missing, write a labelled template and return False (caller exits)."""
    if ENV.exists():
        return True
    ENV.write_text(ENV_TEMPLATE, encoding="utf-8")
    print(f"Created {ENV} — please fill in your details (name, IBAN, car reg, home city,")
    print("signature path), then run matkalasku again. Your .env is gitignored; nothing")
    print("personal is ever shared.")
    return False


def load_env() -> dict:
    cfg: dict = {}
    if not ENV.exists():
        return cfg
    for line in ENV.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, _, v = line.partition("=")
        cfg[k.strip()] = v.strip()
    return cfg


# ── rates (per year, updatable) ──

def load_rates() -> dict:
    try:
        rates = json.loads(RATES.read_text(encoding="utf-8"))
    except Exception:
        rates = {}
    if not rates:
        rates = dict(DEFAULT_RATES)
        try:
            RATES.write_text(json.dumps(rates, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass
    return rates


def rate_for_year(year):
    v = load_rates().get(str(year))
    return float(v) if v is not None else None


def save_rate(year, rate):
    rates = load_rates()
    rates[str(year)] = round(float(rate), 4)
    RATES.write_text(json.dumps(rates, ensure_ascii=False, indent=2), encoding="utf-8")


def latest_known_rate():
    rates = load_rates()
    if not rates:
        return None
    newest = max(rates, key=str)
    return float(rates[newest]), newest


# ── routes (remembered distances) ──

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


def _route_key(o, d):
    return f"{canonical_place(o).lower()}|{canonical_place(d).lower()}"


def load_routes() -> dict:
    try:
        return json.loads(ROUTES.read_text(encoding="utf-8"))
    except Exception:
        return {}


def lookup_cached_km(o, d):
    routes = load_routes()
    for key in (_route_key(o, d), _route_key(d, o)):
        if key in routes:
            try:
                return float(routes[key]["oneway_km"])
            except Exception:
                continue
    return None


def remember_route(o, d, oneway_km, source="stated"):
    routes = load_routes()
    key = _route_key(o, d)
    existing = routes.get(key)
    if existing and source.startswith("auto") and not str(existing.get("source", "")).startswith("auto"):
        return
    routes[key] = {"origin": canonical_place(o), "destination": canonical_place(d),
                   "oneway_km": round(float(oneway_km), 1), "source": source}
    ROUTES.write_text(json.dumps(routes, ensure_ascii=False, indent=2), encoding="utf-8")


def lookup_oneway_km(origin, destination, timeout=8.0):
    """Best-effort one-way driving km via public Nominatim + OSRM (network, not magic)."""
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


def route_string(origin, destination, round_trip=True):
    o, d = origin.strip(), destination.strip()
    return f"{o}–{d}–{o}" if round_trip else f"{o}–{d}"


def format_iban(iban: str) -> str:
    s = re.sub(r"\s+", "", iban or "")
    return " ".join(s[i:i + 4] for i in range(0, len(s), 4)) if s else ""


def prepare_signature(path: Path):
    """Return a transparent-background PNG path (white→alpha, cached) or None."""
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

    def cells(self):
        if len(self.dates) > MAX_TRIPS:
            raise ValueError(f"{len(self.dates)} dates exceeds the template's {MAX_TRIPS} rows")
        if len(self.perdiem) > MAX_PERDIEM:
            raise ValueError(f"{len(self.perdiem)} per-diem rows exceeds {MAX_PERDIEM}")
        cells = {CELL_NAME: self.name, CELL_IBAN: self.iban}
        if self.invoice_date:
            cells[CELL_INVOICE_DATE] = self.invoice_date
        if self.name:
            for c in CELL_NAME_CLARIFY:
                cells[c] = self.name
        for c in BRANDING_CELLS:
            cells[c] = None
        for r in range(FIRST_ROW, LAST_ROW + 1):
            for col in "ABCDEF":
                cells[f"{col}{r}"] = None
        for i, d in enumerate(self.dates):
            r = FIRST_ROW + i
            cells[f"A{r}"] = d
            cells[f"B{r}"] = self.purpose
            cells[f"C{r}"] = self.leg_route(i)
            cells[f"D{r}"] = self.regnr
            cells[f"E{r}"] = self.km_per_leg
            cells[f"F{r}"] = KM_TYPE
        for r in range(PERDIEM_FIRST_ROW, PERDIEM_LAST_ROW + 1):
            for col in ("A", "C", "E", "F"):
                cells[f"{col}{r}"] = None
        for i, laatu in enumerate(self.perdiem):
            r = PERDIEM_FIRST_ROW + i
            cells[f"A{r}"] = self.date_span
            cells[f"C{r}"] = self.destination
            cells[f"E{r}"] = self.purpose
            cells[f"F{r}"] = laatu
        return cells


def fill(data: Matkalasku, out_path, signature: Path | None = None) -> Path:
    area = PRINT_AREA_WITH_PERDIEM if data.perdiem else PRINT_AREA_KM
    sig = None
    sig_file = prepare_signature(signature) if signature else None
    if sig_file:
        sig = (str(sig_file), SIG_ANCHOR_COL, SIG_ANCHOR_ROW, SIG_CX, SIG_CY, SIG_ROW_OFF)
    return xlsx_fill.set_cells(
        TEMPLATE, out_path, SHEET, data.cells(),
        hide_sheets=[FOREIGN_SHEET], print_area=(SHEET, area),
        fit_to_page=True, strip_footer=True, remove_drawing=True, signature=sig,
        extra_cells={FOREIGN_SHEET: {RATE_CELL: data.rate}},
    )


# ── CLI ──

def _ask(prompt, default=""):
    suffix = f" [{default}]" if default else ""
    try:
        ans = input(f"{prompt}{suffix}: ").strip()
    except EOFError:
        ans = ""
    return ans or default


def _slug(s, fallback):
    s = (s or "").translate(str.maketrans("äöåÄÖÅ", "aoaAOA"))
    return re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_").lower() or fallback


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Fill, sign and bake a Finnish Matkalasku PDF.")
    ap.add_argument("--name"); ap.add_argument("--iban"); ap.add_argument("--regnr")
    ap.add_argument("--purpose"); ap.add_argument("--from", dest="origin")
    ap.add_argument("--to", dest="destination")
    ap.add_argument("--km", type=float); ap.add_argument("--km-oneway", dest="km_oneway", type=float)
    ap.add_argument("--date", action="append"); ap.add_argument("--dates")
    ap.add_argument("--place"); ap.add_argument("--invoice-date", dest="invoice_date")
    ap.add_argument("--rate", type=float)
    ap.add_argument("--perdiem", action="append",
                    help="TYPE[:COUNT], TYPE ∈ koko|osa|koko-2|osa-1|ateria")
    ap.add_argument("--auto-km", action="store_true", dest="auto_km")
    ap.add_argument("--out"); ap.add_argument("--no-open", action="store_true", dest="no_open")
    ap.add_argument("--yes", action="store_true")
    args = ap.parse_args(argv)

    if not ensure_env():
        return 0
    cfg = load_env()
    interactive = sys.stdin.isatty() and not args.yes

    name = args.name or (_ask("Nimi (your name)", cfg.get("NAME", "")) if interactive else cfg.get("NAME", ""))
    iban = args.iban or (_ask("Tilinumero (IBAN)", cfg.get("IBAN", "")) if interactive else cfg.get("IBAN", ""))
    purpose = args.purpose or (_ask("Matkan tarkoitus", "Esiintyminen") if interactive else "Esiintyminen")
    regnr = args.regnr or (_ask("Rek.nro (car reg)", cfg.get("CAR_REG", "")) if interactive else cfg.get("CAR_REG", ""))
    home = cfg.get("HOME_CITY", "Helsinki")
    origin = args.origin or (_ask("Lähtöpaikka (from)", home) if interactive else home)
    destination = args.destination or (_ask("Kohde (to)") if interactive else "")
    if not destination:
        print("matkalasku: a destination is required (--to)", file=sys.stderr)
        return 1

    # distance (one-way / per-leg)
    oneway = None
    flag_km = args.km if args.km is not None else args.km_oneway
    if flag_km is not None:
        oneway, route_source = float(flag_km), "stated (flag)"
    else:
        cached = lookup_cached_km(origin, destination)
        auto = None
        if cached is not None:
            print(f"  ↳ remembered route {origin}↔{destination}: {cached:g} km one way")
        if args.auto_km or (interactive and cached is None):
            auto = lookup_oneway_km(origin, destination)
            if auto:
                print(f"  ↳ looked up {origin}→{destination}: ~{auto} km one way")
        default = cached if cached is not None else auto
        route_source = "stated"
        if interactive:
            while oneway is None:
                ow = _ask("Yhdensuuntainen matka km (one-way, per leg)",
                          f"{default:g}" if default is not None else "")
                if not ow:
                    if default is not None:
                        oneway = float(default)
                    continue
                try:
                    oneway = float(ow.replace(",", "."))
                except ValueError:
                    print("    (please type a number, e.g. 189)")
            route_source = "remembered" if (cached is not None and oneway == cached) else "stated"
        elif default is not None:
            oneway = float(default)
            route_source = "remembered" if cached is not None else "auto (osrm)"
        else:
            print("matkalasku: a distance is required (--km / --auto-km)", file=sys.stderr)
            return 1
    if oneway and oneway > 0:
        remember_route(origin, destination, oneway, source=route_source)

    # dates
    raw_dates = list(args.date or [])
    if args.dates:
        raw_dates += [p for p in re.split(r"[,\s]+", args.dates) if p]
    if not raw_dates and interactive:
        try:
            n = max(1, int(_ask("Montako matkapäivää (travel dates)", "1")))
        except ValueError:
            n = 1
        for i in range(n):
            d = _ask(f"  Päiväys {i + 1} (YYYY-MM-DD)")
            if d:
                raw_dates.append(d)
    if not raw_dates:
        print("matkalasku: at least one travel date is required (--date)", file=sys.stderr)
        return 1
    dates = [normalize_date(d) for d in raw_dates]

    # per-diem (optional)
    perdiem = []
    for spec in (args.perdiem or []):
        typ, _, cnt = spec.partition(":")
        nn = int(cnt) if cnt.strip().isdigit() else 1
        perdiem += [PERDIEM_TYPES.get(typ.strip().lower(), typ.strip())] * nn

    # rate for the travel year (stored, updatable)
    year = dates[0].split(".")[-1]
    if args.rate is not None:
        rate = float(args.rate); save_rate(year, rate)
    else:
        rate = rate_for_year(year)
        if rate is not None:
            print(f"  ↳ {year} kilometrikorvaus: {rate:g} €/km")
        elif interactive:
            fb = latest_known_rate()
            ans = _ask(f"Vuoden {year} kilometrikorvaus €/km (not on record)", f"{fb[0]:g}" if fb else "")
            rate = float(ans.replace(",", ".")) if ans else (fb[0] if fb else RATE_FALLBACK)
            save_rate(year, rate)
            print(f"  ↳ stored {year} = {rate:g} €/km")
        else:
            fb = latest_known_rate()
            rate = fb[0] if fb else RATE_FALLBACK
            print(f"  ↳ no rate for {year}; using {rate:g} €/km (set with --rate)")

    place_label = args.place or (_ask("Paikka / tapahtuma (place/event)", destination) if interactive else destination)
    invoice_date = normalize_date(args.invoice_date) if args.invoice_date else datetime.now().strftime("%-d.%-m.%Y")

    data = Matkalasku(name=name, iban=format_iban(iban) or iban, purpose=purpose, regnr=regnr,
                      origin=origin, destination=destination, km_per_leg=oneway, dates=dates,
                      perdiem=perdiem, rate=rate, invoice_date=invoice_date)

    handle = cfg.get("HANDLE") or _slug(name, "matkalasku")
    place = _slug(place_label, "matka")
    dd, mm, yyyy = (dates[0].split(".") + ["", "", year])[:3]
    date_part = f"{yyyy}_{int(mm):02d}_{int(dd):02d}" if mm.isdigit() and dd.isdigit() else f"{year}_01_01"
    out = Path(args.out) if args.out else (OUT_DIR / f"{date_part}_{handle}_matkalasku_{place}.xlsx")
    out.parent.mkdir(parents=True, exist_ok=True)

    sig_path = Path(cfg.get("SIGNATURE_PATH", "")) if cfg.get("SIGNATURE_PATH") else None
    if sig_path and not sig_path.is_absolute():
        sig_path = REPO / sig_path
    try:
        fill(data, out, signature=sig_path)
    except Exception as e:  # noqa: BLE001
        print(f"matkalasku: could not fill template: {e}", file=sys.stderr)
        return 1
    pdf = to_pdf(out, out.parent)
    record = {**{k: getattr(data, k) for k in
                 ("name", "iban", "regnr", "purpose", "origin", "destination",
                  "km_per_leg", "dates", "perdiem", "total_km", "rate", "total_eur", "invoice_date")},
              "route": data.route, "xlsx": str(out), "pdf": str(pdf) if pdf else None}
    out.with_suffix(".json").write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"\n✓ Matkalasku — {data.name or '(no name)'} · {data.route}")
    for i, d in enumerate(data.dates):
        print(f"    {d}  {data.leg_route(i)}  {data.km_per_leg:g} km")
    print(f"  {len(data.dates)} leg(s) × {data.km_per_leg:g} km = {data.total_km:g} km "
          f"× {data.rate:g} €/km = {data.total_eur:.2f} €")
    print(f"  Päiväys: {data.invoice_date}   {'✍ signed' if sig_path and sig_path.exists() else '(no signature)'}")
    print(f"  xlsx → {out}")
    print(f"  pdf  → {pdf if pdf else '(install LibreOffice to bake the PDF)'}")
    if pdf and not args.no_open and sys.platform == "darwin":
        import subprocess
        subprocess.call(["open", str(pdf)])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

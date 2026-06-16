#!/usr/bin/env python3
"""matkalasku — fill, sign and bake a Finnish travel-expense claim (Matkalasku) as a PDF.

This is the `.env` front-end: it reads your details from a local (gitignored) `.env`,
asks where/when you drove, and hands off to matkalasku_core (the shared engine). No
cloud, no account — nothing personal is ever shared.

First run creates a labelled `.env` for you to fill in. Then: `python matkalasku.py`
(interactive) or with flags (see --help).

Deps: standard library for the core. Optional: Pillow (transparent signature) and
LibreOffice (`soffice`, bakes the PDF — without it you get the .xlsx).
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import matkalasku_core as core

REPO = Path(__file__).resolve().parent
ENV = REPO / ".env"
RATES = REPO / "rates.json"
ROUTES = REPO / ".routes.json"
TEMPLATE = REPO / "templates" / "matkalasku-2026.xlsx"
OUT_DIR = REPO / "out"

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


def ensure_env() -> bool:
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
    ap.add_argument("--template", help="path to a Matkalasku .xlsx (default: bundled 2026)")
    ap.add_argument("--no-verify", action="store_true", dest="no_verify",
                    help="skip the template/profile self-check (not recommended)")
    ap.add_argument("--out"); ap.add_argument("--no-open", action="store_true", dest="no_open")
    ap.add_argument("--yes", action="store_true")
    args = ap.parse_args(argv)

    if not ensure_env():
        return 0
    cfg = load_env()
    interactive = sys.stdin.isatty() and not args.yes
    template = Path(args.template) if args.template else TEMPLATE
    if not template.exists():
        print(f"matkalasku: template not found: {template}", file=sys.stderr)
        return 1
    profile = core.load_template_profile(template)

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
        cached = core.lookup_cached_km(ROUTES, origin, destination)
        auto = None
        if cached is not None:
            print(f"  ↳ remembered route {origin}↔{destination}: {cached:g} km one way")
        if args.auto_km or (interactive and cached is None):
            auto = core.lookup_oneway_km(origin, destination)
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
        core.remember_route(ROUTES, origin, destination, oneway, source=route_source)

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
    dates = [core.normalize_date(d) for d in raw_dates]

    # per-diem (optional): flags win; otherwise ask interactively (off by default)
    perdiem = []
    perdiem_types = (profile.get("perdiem") or {}).get("types", {})
    for spec in (args.perdiem or []):
        typ, _, cnt = spec.partition(":")
        nn = int(cnt) if cnt.strip().isdigit() else 1
        perdiem += [perdiem_types.get(typ.strip().lower(), typ.strip())] * nn
    if not perdiem and interactive:
        perdiem = core.prompt_perdiem(_ask, perdiem_types)

    # rate for the travel year (stored, updatable)
    year = dates[0].split(".")[-1]
    if args.rate is not None:
        rate = float(args.rate); core.save_rate(RATES, year, rate)
    else:
        rate = core.rate_for_year(RATES, year)
        if rate is not None:
            print(f"  ↳ {year} kilometrikorvaus: {rate:g} €/km")
        elif interactive:
            fb = core.latest_known_rate(RATES)
            ans = _ask(f"Vuoden {year} kilometrikorvaus €/km (not on record)", f"{fb[0]:g}" if fb else "")
            rate = float(ans.replace(",", ".")) if ans else (fb[0] if fb else core.RATE_FALLBACK)
            core.save_rate(RATES, year, rate)
            print(f"  ↳ stored {year} = {rate:g} €/km")
        else:
            fb = core.latest_known_rate(RATES)
            rate = fb[0] if fb else core.RATE_FALLBACK
            print(f"  ↳ no rate for {year}; using {rate:g} €/km (set with --rate)")

    place_label = args.place or (_ask("Paikka / tapahtuma (place/event)", destination) if interactive else destination)
    invoice_date = core.normalize_date(args.invoice_date) if args.invoice_date else core.fi_date(datetime.now())

    data = core.Matkalasku(name=name, iban=core.format_iban(iban) or iban, purpose=purpose,
                           regnr=regnr, origin=origin, destination=destination,
                           km_per_leg=oneway, dates=dates, perdiem=perdiem, rate=rate,
                           invoice_date=invoice_date)

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
        core.fill(data, out, profile, template, signature=sig_path, verify=not args.no_verify)
    except Exception as e:  # noqa: BLE001
        print(f"matkalasku: {e}", file=sys.stderr)
        return 1
    pdf = core.to_pdf(out, out.parent)
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

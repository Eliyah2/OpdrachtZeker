#!/usr/bin/env python3
"""Bouwt sites/public/ — de statische bestanden die naar Freebuff Sites gaan.

Bron van waarheid blijft landingspagina/index.html en risicoscan/index.html.
Dit script kopieert die ongewijzigd naar de deploymap, zodat er geen tweede,
met de hand aangepaste versie kan ontstaan.

Gebruik:
    python scripts/bouw_site.py           # (her)bouw sites/public
    python scripts/bouw_site.py --check   # faalt als de deploymap niet synchroon is
"""

from __future__ import annotations

import sys
from pathlib import Path

WORTEL = Path(__file__).resolve().parent.parent
DOEL = WORTEL / "sites" / "public"

# (bron, doelpad binnen sites/public)
BESTANDEN = [
    ("landingspagina/index.html", "index.html"),
    ("landingspagina/index.html", "landingspagina/index.html"),
    ("risicoscan/index.html", "risicoscan/index.html"),
]


def bouw() -> int:
    gewijzigd = 0
    for bron, doel in BESTANDEN:
        bronpad = WORTEL / bron
        doelpad = DOEL / doel
        if not bronpad.is_file():
            print(f"FOUT: bronbestand ontbreekt: {bron}")
            return 1
        inhoud = bronpad.read_bytes()
        if doelpad.is_file() and doelpad.read_bytes() == inhoud:
            continue
        doelpad.parent.mkdir(parents=True, exist_ok=True)
        doelpad.write_bytes(inhoud)
        print(f"  geschreven: sites/public/{doel}")
        gewijzigd += 1
    if gewijzigd == 0:
        print("sites/public is al synchroon met de bronbestanden.")
    else:
        print(f"klaar: {gewijzigd} bestand(en) bijgewerkt.")
    return 0


def controleer() -> int:
    problemen = []
    for bron, doel in BESTANDEN:
        bronpad = WORTEL / bron
        doelpad = DOEL / doel
        if not doelpad.is_file():
            problemen.append(f"sites/public/{doel} ontbreekt")
        elif doelpad.read_bytes() != bronpad.read_bytes():
            problemen.append(f"sites/public/{doel} loopt uit de pas met {bron}")
    if problemen:
        print("FOUT: de deploymap is niet synchroon met de bronbestanden:")
        for p in problemen:
            print(f"  - {p}")
        print("Los dit op met: python scripts/bouw_site.py")
        return 1
    print("site-build OK: sites/public is synchroon met landingspagina/ en risicoscan/.")
    return 0


if __name__ == "__main__":
    sys.exit(controleer() if "--check" in sys.argv else bouw())

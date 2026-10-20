"""
Pobiera aktualny Rejestr Srodkow Ochrony Roslin ze strony Ministerstwa Rolnictwa
i zapisuje go jako rejestr.json w formacie, ktory czyta aplikacja.

Uruchomienie normalne:   python aktualizuj_rejestr.py
Test na lokalnym pliku:  python aktualizuj_rejestr.py --plik rejestr_sor_20261016.xlsx
"""
import json
import os
import re
import sys
from urllib.parse import urljoin

import pandas as pd

STRONA = "https://www.gov.pl/web/rolnictwo/rejestr-rodkow-ochrony-roslin"
WZORZEC_PLIKU = re.compile(r"rejestr_sor_(\d{8})\.xlsx", re.IGNORECASE)
PLIK_WYJSCIOWY = "rejestr.json"
MIN_SRODKOW = 2500          # zabezpieczenie: mniej oznacza, ze cos poszlo nie tak
OCZEKIWANA_LICZBA_KOLUMN = 17
NAGLOWKI_HTTP = {"User-Agent": "Mozilla/5.0 (aktualizacja rejestru dla aplikacji rolnika)"}

KOLUMNY = ['lp', 'nr', 'nazwa', 'rodzaj', 'podmiot', 'producent', 'substancja', 'zdrowie',
           'pszczoly', 'wodne', 'srodowisko', 'wazne_do', 'zapasy_sprzedaz',
           'zapasy_stosowanie', 'opakowania', 'referencyjny', 'uwagi']


def znajdz_link(html):
    """Szuka na stronie linku do polskiego rejestru (nie angielskiego)."""
    from bs4 import BeautifulSoup
    zupa = BeautifulSoup(html, "html.parser")
    for a in zupa.find_all("a", href=True):
        tekst = a.get_text(" ", strip=True).replace("\u200b", "")
        m = WZORZEC_PLIKU.search(tekst)
        if m:
            return urljoin(STRONA, a["href"]), m.group(1)
    return None, None


def substancje(tekst):
    if not isinstance(tekst, str):
        return []
    tekst = tekst.replace('_x000D_', '')
    wynik = []
    for linia in tekst.split('\n'):
        linia = linia.strip().rstrip(',').strip()
        if not linia:
            continue
        nazwa = linia.rsplit(' - ', 1)[0].strip() if ' - ' in linia else linia
        if nazwa and nazwa.lower() not in [w.lower() for w in wynik]:
            wynik.append(nazwa)
    return wynik


def data_iso(x):
    return x.strftime('%Y-%m-%d') if pd.notna(x) else ''


def konwertuj(zrodlo, data_rrrrmmdd):
    df = pd.read_excel(zrodlo, sheet_name=0)
    if len(df.columns) != OCZEKIWANA_LICZBA_KOLUMN:
        raise SystemExit(
            f"BLAD: rejestr ma {len(df.columns)} kolumn zamiast {OCZEKIWANA_LICZBA_KOLUMN}. "
            "Ministerstwo zmienilo uklad pliku, trzeba dostosowac skrypt."
        )
    df.columns = KOLUMNY

    srodki = []
    for _, r in df.iterrows():
        if not isinstance(r.nazwa, str) or not r.nazwa.strip():
            continue
        srodki.append({
            'n': r.nazwa.strip(),
            'r': str(r.rodzaj).strip() if pd.notna(r.rodzaj) else '',
            's': substancje(r.substancja),
            'w': data_iso(r.wazne_do),
            'z': data_iso(r.zapasy_stosowanie),
        })

    if len(srodki) < MIN_SRODKOW:
        raise SystemExit(f"BLAD: tylko {len(srodki)} srodkow (minimum {MIN_SRODKOW}). Nie zapisuje.")
    z_data = sum(1 for s in srodki if s['w'])
    z_subst = sum(1 for s in srodki if s['s'])
    if z_data < 0.98 * len(srodki) or z_subst < 0.98 * len(srodki):
        raise SystemExit("BLAD: za duzo srodkow bez terminu waznosci lub substancji. Nie zapisuje.")

    d = data_rrrrmmdd
    return {'data': f"{d[6:8]}.{d[4:6]}.{d[0:4]}", 'srodki': srodki}


def data_z_tekstu(tekst):
    """dd.mm.rrrr -> rrrrmmdd (do porownywania wersji)"""
    return tekst[6:10] + tekst[3:5] + tekst[0:2]


def main():
    if len(sys.argv) >= 3 and sys.argv[1] == "--plik":
        sciezka = sys.argv[2]
        m = WZORZEC_PLIKU.search(os.path.basename(sciezka))
        if not m:
            raise SystemExit("Nazwa pliku musi miec postac rejestr_sor_RRRRMMDD.xlsx")
        wynik = konwertuj(sciezka, m.group(1))
    else:
        import io
        import requests
        strona = requests.get(STRONA, headers=NAGLOWKI_HTTP, timeout=60)
        strona.raise_for_status()
        link, data = znajdz_link(strona.text)
        if not link:
            raise SystemExit("BLAD: nie znaleziono na stronie linku do pliku rejestr_sor_*.xlsx.")
        print("Znaleziono:", link, data)

        # Nie cofamy sie do starszej wersji
        if os.path.exists(PLIK_WYJSCIOWY):
            with open(PLIK_WYJSCIOWY, encoding="utf-8") as f:
                obecna = json.load(f).get("data", "")
            if obecna and data_z_tekstu(obecna) >= data:
                print(f"Rejestr z {obecna} jest aktualny. Nic do zrobienia.")
                return

        plik = requests.get(link, headers=NAGLOWKI_HTTP, timeout=120)
        plik.raise_for_status()
        if not plik.content.startswith(b"PK"):
            raise SystemExit("BLAD: pobrany plik nie wyglada na plik Excel (.xlsx).")
        wynik = konwertuj(io.BytesIO(plik.content), data)

    with open(PLIK_WYJSCIOWY, "w", encoding="utf-8") as f:
        json.dump(wynik, f, ensure_ascii=False, separators=(',', ':'))
    print(f"Zapisano {len(wynik['srodki'])} srodkow, wersja z {wynik['data']}.")


if __name__ == "__main__":
    main()

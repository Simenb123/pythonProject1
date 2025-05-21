# Bilagsverktøy

Dette prosjektet inneholder noen enkle Python-skript for å velge klient og trekke ut bilag.

## Avhengigheter

- Python 3 med `tkinter` aktivert (for GUI)
- [`pandas`](https://pandas.pydata.org/)
- [`openpyxl`](https://openpyxl.readthedocs.io/)
- [`chardet`](https://pypi.org/project/chardet/)
- [`pyarrow`](https://arrow.apache.org/docs/python/) – benyttes for å skrive parquet-filer

Installer avhengighetene med `pip`:

```bash
pip install pandas openpyxl chardet pyarrow
```

## Kjøre `klientvelger.py`

Programmet lar deg velge en klientmappe og starte bilags‑GUI. Kjør med Python:

```bash
python klientvelger.py
```

Første gang må du endre `ROOT_DIR` i `klientvelger.py` slik at det peker på mappen der klientene dine ligger. Programmet husker sist valgte klient mellom hver gang det kjøres.

Etter at du har valgt klient og bilagsfil starter `bilag_gui_tk.py` automatisk.

## Kjøre `bilag_gui_tk.py` direkte

`bilag_gui_tk.py` kan også startes manuelt dersom du gir en bilagsfil som argument:

```bash
python bilag_gui_tk.py /path/til/bilag.xlsx
```

Filen kan være i Excel- eller CSV-format. Programmet forsøker å lese eventuelle `_mapping.json`-filer for å vite hvilke kolonner som inneholder kontonummer, beløp, dato og bilagsnummer.

### Live-statistikk

Mens du justerer konto‑ og beløpsintervall vises en liten tekstlinje under feltene som oppdateres hvert 0,3 sekund. Den viser antall linjer i intervallet, summen av beløp og snittbeløpet. Statistikken fungerer så snart bilagsfilen er lest inn og nødvendige kolonner er identifisert.

Når du trykker **Trekk utvalg** blir et tilfeldig utvalg av bilag valgt, skrevet til en Excel‑fil og du får beskjed om hvor filen ligger.

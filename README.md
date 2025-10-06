# Bilagsverktøy

Bilagsverktøy er en samling Python-verktøy for å jobbe med klientdata og
bilagsfiler. Prosjektet kombinerer frittstående skript i rotmappen med et
omfattende GUI- og tjenestelag under `src/app/`.

## Hovedkomponenter

- **Klientvelger** (`klientvelger.py`): et lite Tkinter-program som lar deg velge
  klientmappe, datakilde og bilagsfil før du starter analyse- eller
  uttrekksflyten. Programmet husker sist valgte klient og lagrer metadata per
  klient.
- **Bilags-GUI** (`bilag_gui_tk.py`): hovedvinduet for bilagsanalyse og
  uttrekk. Leser Excel/CSV-filer, foreslår kolonnemapping og viser løpende
  statistikk for filtrerte utvalg.
- **Klientportal** (`src/app/gui/start_portal.py`): et større startvindu som
  samler klientadministrasjon, teamhåndtering og inngang til
  Klienthub-/analyseflyten. Portal og tilhørende GUI-moduler ligger under
  `src/app/gui/`.
- **Tjenestelag** (`src/app/services/`): funksjoner for blant annet
  klientregister, importpipen, mapping, versjonshåndtering og regnskapslinjer.
  GUI-ene bygger videre på disse modulene.

I tillegg finnes flere spesialiserte skript i rotmappen (f.eks.
`run_fifo.py`, `import_pipeline.py`, `convert_maestro_sb123.py`) for
engangsoppgaver og integrasjoner.

## Kodeoppsett

```
├── klientvelger.py              # Hurtigstart av bilagsflyten
├── bilag_gui_tk.py              # Frittstående bilags-GUI
├── src/
│   └── app/
│       ├── gui/                 # Tkinter-apper (klientportal, hub, import-GUIer, widgets)
│       ├── services/            # Klientregister, import, mapping, versjonering m.m.
│       ├── parsers/             # SAF-T/GL-parsing og andre lesere
│       ├── mva/                 # MVA-relaterte hjelpere
│       ├── converters/          # Konverteringsskript for eksterne formater
│       ├── dokumentreader/      # Lesing av dokumentdata
│       ├── glscanner/           # Analyse av hovedbok/saldobalanse
│       └── a07/, aksjonærregister/ ... # Domenespesifikke moduler
└── tests/                       # Pytest-baserte tester for tjenester og kontrollere
```

Se `src/app/gui/` for flere brukergrensesnitt, bl.a. Klienthub,
master-import, aksjonærregister-import og teamredigering.
Tjenestelagene i `src/app/services/` og parserne under `src/app/parsers/`
leverer forretningslogikken som GUI-ene bygger på.

## Krav og installasjon

Prosjektet bruker Python 3.10+ med Tkinter-støtte samt eksterne biblioteker som
`pandas`, `openpyxl`, `xlrd`, `xlwings` og `tkcalendar`. Installér avhengigheter
via `requirements.txt`:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Kjøreeksempler

Start tradisjonell klientvelger:

```bash
python klientvelger.py
```

Start klientportalen (inkluderer Klienthub og administrative verktøy):

```bash
python -m app.gui.start_portal
```

GUI-ene bruker Tkinter; på Linux må du ha et Python-bygg med `tk` installert.

## Utvikling

Kjør testene med pytest:

```bash
python -m pytest
```

Når du jobber med GUI-modulene kan det være nyttig å sette `PYTHONPATH` til
`src/` slik at modulimportene fungerer når du kjører filer direkte:

```bash
PYTHONPATH=src python src/app/gui/start_portal.py
```


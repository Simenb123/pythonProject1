# -*- coding: utf-8 -*-
"""
MVA-termin-dashboard fra SAF-T (NO)
-----------------------------------
- Leser SAF-T Regnskap (XML)
- Aggregerer grunnlag og mva per SAF-T mva-kode og per termin (bimånedlig som Skatteetaten)
- Viser GUI (Streamlit) med:
  * filtrering per år
  * tabeller på kode- og termin-nivå
  * nedlasting av CSV
  * frivillig sammenligning mot "faktisk innrapportert" (opplasting av egen CSV)
- Støtter opplasting av egen mapping (CSV) for å klassifisere mva-koder som Utgående/Inngående/Annet.

Avhengigheter: streamlit, pandas, lxml, python-dateutil
"""

from __future__ import annotations

import io
import math
from dataclasses import dataclass
from datetime import datetime
from dateutil import parser as dtparser
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
import streamlit as st
from lxml import etree

# ----- Konstanter og standardinnstillinger ---------------------------------

# SAF-T mva-koder som normalt ikke rapporteres i mva-meldingen (filtreres bort)
# Kilde: Skatteetatens informasjonsmodell for MVA-meldingen (koder 0, 7, 20, 21, 22).
# https://skatteetaten.github.io/mva-meldingen/mvameldingen/informasjonsmodell/
NON_REPORTABLE_SAFT_CODES = {"0", "7", "20", "21", "22"}

# Standard mapping av SAF-T mva-koder til "retning"/kategori.
# Du kan overstyre i appen ved å laste opp egen CSV (se format under).
# Kilde for beskrivelse av koder: Regnskap Norge-tabell (se dokumentasjon).
DEFAULT_TAXCODE_MAP = pd.DataFrame(
    [
        # Utgående (avgiftspliktig omsetning)
        {"TaxCode": "3",  "Direction": "UTG", "Label": "Salg/uttak – høy sats"},
        {"TaxCode": "31", "Direction": "UTG", "Label": "Salg/uttak – middels sats"},
        {"TaxCode": "32", "Direction": "UTG", "Label": "Salg fisk/marine ressurser"},
        {"TaxCode": "33", "Direction": "UTG", "Label": "Salg/uttak – lav sats"},

        # Fritak/Unntak (grunnlag rapporteres – mva typisk 0)
        {"TaxCode": "5",  "Direction": "FRITAK_UNNTAK", "Label": "Fritatt merverdiavgift"},
        {"TaxCode": "52", "Direction": "FRITAK_UNNTAK", "Label": "Klimakvoter/gull – fritak"},
        {"TaxCode": "6",  "Direction": "FRITAK_UNNTAK", "Label": "Unntatt mvaloven"},

        # Innførsel av varer (grunnlagskoder; mva beregnes normalt via Tolletaten)
        {"TaxCode": "81", "Direction": "INFØRSEL", "Label": "Innførsel m/fradragsrett (høy)"},
        {"TaxCode": "82", "Direction": "INFØRSEL", "Label": "Innførsel u/fradragsrett (høy)"},
        {"TaxCode": "83", "Direction": "INFØRSEL", "Label": "Innførsel m/fradragsrett (middels)"},
        {"TaxCode": "84", "Direction": "INFØRSEL", "Label": "Innførsel u/fradragsrett (middels)"},
        {"TaxCode": "85", "Direction": "INFØRSEL", "Label": "Innførsel – nullsats"},

        # Fjernleverbare tjenester (omvendt avgiftsplikt – beregnet utgående)
        {"TaxCode": "86", "Direction": "RC_BEREGNET_UTG", "Label": "Kjøp tjenester fra utland – RC (høy)"},
        {"TaxCode": "87", "Direction": "RC_BEREGNET_UTG", "Label": "Kjøp tjenester fra utland – RC u/fradrag (høy)"},
        {"TaxCode": "88", "Direction": "RC_BEREGNET_UTG", "Label": "Kjøp tjenester fra utland – RC (lav)"},
        {"TaxCode": "89", "Direction": "RC_BEREGNET_UTG", "Label": "Kjøp tjenester fra utland – RC u/fradrag (lav)"},

        # Klimakvoter/gull (kjøp)
        {"TaxCode": "91", "Direction": "ING", "Label": "Kjøp klimakvoter/gull m/fradrag"},
        {"TaxCode": "92", "Direction": "ING", "Label": "Kjøp klimakvoter/gull u/fradrag"},

        # Inngående mva – innenlands kjøp
        {"TaxCode": "1",  "Direction": "ING", "Label": "Kjøp – høy sats"},
        {"TaxCode": "11", "Direction": "ING", "Label": "Kjøp – middels sats"},
        {"TaxCode": "12", "Direction": "ING", "Label": "Kjøp fisk/marine ressurser"},
        {"TaxCode": "13", "Direction": "ING", "Label": "Kjøp – lav sats"},

        # Inngående mva – fradrag import
        {"TaxCode": "14", "Direction": "ING", "Label": "Fradrag innførselsmva (høy)"},
        {"TaxCode": "15", "Direction": "ING", "Label": "Fradrag innførselsmva (middels)"},
    ]
)

# Felt som appen ser etter under <TaxInformation> i SAF-T:
TAX_BASE_CANDIDATES = ("TaxBase", "TaxBaseAmount", "TaxableAmount", "TaxBasisAmount", "BaseAmount")
TAX_AMOUNT_CANDIDATES = ("TaxAmount", "Amount")

# ------------------------- Hjelpefunksjoner ---------------------------------

def localname(tag: str) -> str:
    """Fjerner XML-namespace fra et tag-navn."""
    if tag is None:
        return ""
    return tag.split("}", 1)[-1] if "}" in tag else tag

def try_parse_date(text: Optional[str]) -> Optional[datetime]:
    if not text:
        return None
    try:
        return dtparser.parse(text).date()
    except Exception:
        return None

def to_float(x: Optional[str]) -> Optional[float]:
    if x is None:
        return None
    try:
        # erstatt ev. tusenskilletegn
        s = str(x).strip().replace(" ", "").replace("\u00A0", "").replace(",", ".")
        return float(s)
    except Exception:
        return None

def first_child_text(elem: etree._Element, names: Iterable[str]) -> Optional[str]:
    """Returner første eksisterende barneelementtekst blant gitt navneliste (uten namespace)."""
    for ch in elem:
        if localname(ch.tag) in names:
            return (ch.text or "").strip()
    return None

def extract_taxinfo(line_elem: etree._Element) -> Optional[Dict[str, Optional[str]]]:
    """Hent TaxInformation fra en Line."""
    for ch in line_elem:
        if localname(ch.tag) == "TaxInformation":
            # Sjekk at det faktisk er VAT
            tax_type = first_child_text(ch, ("TaxType",))
            if tax_type and tax_type.strip().upper() != "VAT":
                return None
            tax_code = first_child_text(ch, ("TaxCode",))
            tax_base = first_child_text(ch, TAX_BASE_CANDIDATES)
            tax_amount = first_child_text(ch, TAX_AMOUNT_CANDIDATES)
            return {
                "TaxCode": (tax_code or "").strip(),
                "TaxBase": tax_base,
                "TaxAmount": tax_amount,
            }
    return None

def month_to_termin(month: int, periodicity: str = "6-terminer") -> Tuple[int, str]:
    """Map måned til termin-nummer og en pen etikett."""
    m = int(month)
    if periodicity == "12-måneder":
        return m, f"M{m:02d}"
    if periodicity == "4-kvartaler":
        q = (m - 1) // 3 + 1
        return q, f"K{q}"
    if periodicity == "1-år":
        return 1, "År"
    # Standard: 6 terminer
    termin = (m + 1) // 2
    labels = {1: "T1 (Jan–Feb)", 2: "T2 (Mar–Apr)", 3: "T3 (Mai–Jun)",
              4: "T4 (Jul–Aug)", 5: "T5 (Sep–Okt)", 6: "T6 (Nov–Des)"}
    return termin, labels.get(termin, f"T{termin}")

# -------------------- SAF-T parser (streaming, minnevennlig) ----------------

def parse_saft_xml(file_like) -> pd.DataFrame:
    """
    Returnerer DataFrame med kolonnene:
    ['Date', 'Year', 'Month', 'Termin', 'TerminLabel', 'TaxCode', 'TaxBase', 'TaxAmount']
    Kun linjer med TaxInformation (VAT) tas med.
    """
    rows: List[Dict] = []
    context = etree.iterparse(file_like, events=("start", "end"))
    current_date: Optional[datetime] = None
    current_month: Optional[int] = None

    for event, elem in context:
        tag = localname(elem.tag)

        if event == "start":
            if tag == "JournalEntry":
                current_date = None
                current_month = None

        elif event == "end":
            # Finn transaksjonsdato på nivå JournalEntry -> TransactionDate
            if tag == "TransactionDate":
                current_date = try_parse_date(elem.text)
                if current_date:
                    current_month = current_date.month

            elif tag == "Line":
                tx = extract_taxinfo(elem)
                if tx and tx["TaxCode"]:
                    tc = tx["TaxCode"]
                    if tc not in NON_REPORTABLE_SAFT_CODES:
                        # fall-back: om TransactionDate mangler, prøv å finne dato på linjen
                        date_for_line = current_date
                        if date_for_line is None:
                            # noen systemer legger 'PostingDate' eller lignende på linje/dokument
                            maybe_dates = []
                            for ch in elem:
                                nm = localname(ch.tag)
                                if nm in {"PostingDate", "PostingDateTime", "Date"}:
                                    maybe_dates.append(ch.text)
                            for text in maybe_dates:
                                date_for_line = try_parse_date(text)
                                if date_for_line:
                                    break

                        if date_for_line is None:
                            # Uten dato får vi ikke termin – hopp over
                            elem.clear()
                            continue

                        year = date_for_line.year
                        month = date_for_line.month
                        termin, label = month_to_termin(month, periodicity="6-terminer")

                        tax_base_f = to_float(tx["TaxBase"])
                        tax_amount_f = to_float(tx["TaxAmount"])

                        rows.append(
                            {
                                "Date": date_for_line,
                                "Year": year,
                                "Month": month,
                                "Termin": termin,
                                "TerminLabel": label,
                                "TaxCode": tc,
                                "TaxBase": tax_base_f if tax_base_f is not None else 0.0,
                                "TaxAmount": tax_amount_f if tax_amount_f is not None else 0.0,
                            }
                        )
                # Rydd element fra minne
                elem.clear()

            elif tag == "JournalEntry":
                elem.clear()

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    # Normaliser kode som streng uten ledende nuller
    df["TaxCode"] = df["TaxCode"].astype(str).str.strip().str.lstrip("0")
    return df

# ---------------------- Aggregat & sammenligning ----------------------------

def load_mapping_df(uploaded_csv: Optional[io.BytesIO]) -> pd.DataFrame:
    """Les brukerens mapping-fil (CSV) eller returner default mapping."""
    if uploaded_csv is None:
        return DEFAULT_TAXCODE_MAP.copy()
    try:
        df = pd.read_csv(uploaded_csv, dtype={"TaxCode": str})
        # Forventede kolonner: TaxCode, Direction, Label (flere kolonner tillates)
        missing = {"TaxCode", "Direction"} - set(df.columns)
        if missing:
            st.warning(f"Mapping-CSV mangler kolonner: {', '.join(sorted(missing))}. "
                       f"Faller tilbake til standard mapping.")
            return DEFAULT_TAXCODE_MAP.copy()
        df["TaxCode"] = df["TaxCode"].astype(str).str.strip().str.lstrip("0")
        return df
    except Exception as e:
        st.error(f"Kunne ikke lese mapping-CSV: {e}")
        return DEFAULT_TAXCODE_MAP.copy()

def signed_vat(direction: str, tax_amount: float) -> float:
    """Gjør Inngående negativt, Utgående positivt; andre kategorier 0 (for netto-beregning)."""
    if direction == "ING":
        return -float(tax_amount or 0.0)
    if direction in {"UTG", "RC_BEREGNET_UTG"}:
        return float(tax_amount or 0.0)
    return 0.0

def summarize_by_code(df: pd.DataFrame, mapping_df: pd.DataFrame,
                      years: List[int]) -> pd.DataFrame:
    """Summer per år/termin/SAF-T-kode, og slå på mapping."""
    if df.empty:
        return df

    df2 = df[df["Year"].isin(years)].copy()
    mp = mapping_df.copy()
    mp["TaxCode"] = mp["TaxCode"].astype(str).str.strip().str.lstrip("0")

    merged = df2.merge(mp, how="left", on="TaxCode")
    merged["Direction"] = merged["Direction"].fillna("UKJENT")
    merged["Label"] = merged["Label"].fillna("")

    # Signert mva for netto-beregning
    merged["VAT_Signed"] = merged.apply(
        lambda r: signed_vat(str(r["Direction"]), float(r["TaxAmount"])), axis=1
    )

    grp = (
        merged.groupby(
            ["Year", "Termin", "TerminLabel", "TaxCode", "Direction", "Label"],
            dropna=False,
            as_index=False,
        )
        .agg(TaxBase=("TaxBase", "sum"), TaxAmount=("TaxAmount", "sum"), VAT_Signed=("VAT_Signed", "sum"))
    )

    # Runder penere for visning
    for c in ("TaxBase", "TaxAmount", "VAT_Signed"):
        grp[c] = grp[c].round(2)

    return grp.sort_values(["Year", "Termin", "TaxCode"])

def summarize_by_termin(grp_code: pd.DataFrame) -> pd.DataFrame:
    """Summer per år/termin på hovedkategorier og lag netto (UTG - ING)."""
    if grp_code.empty:
        return grp_code

    # summer i to nivåer: per Direction og totalt
    per_dir = (
        grp_code.groupby(["Year", "Termin", "TerminLabel", "Direction"], as_index=False)
        .agg(TaxBase=("TaxBase", "sum"), TaxAmount=("TaxAmount", "sum"), VAT_Signed=("VAT_Signed", "sum"))
    )

    # Pivotér utgående/inngående
    pivot = per_dir.pivot_table(
        index=["Year", "Termin", "TerminLabel"],
        columns="Direction",
        values=["TaxBase", "TaxAmount", "VAT_Signed"],
        aggfunc="sum",
        fill_value=0.0,
    )

    # flatten multiindex columns
    pivot.columns = [f"{a}_{b}" for a, b in pivot.columns]
    pivot = pivot.reset_index()

    # Definer utgående og inngående kolonner som kan finnes
    out_cols = [c for c in pivot.columns if c.endswith("_UTG") or c.endswith("_RC_BEREGNET_UTG")]
    in_cols = [c for c in pivot.columns if c.endswith("_ING")]

    pivot["MVA_Utgående"] = pivot[[c for c in pivot.columns if c == "TaxAmount_UTG" or c == "TaxAmount_RC_BEREGNET_UTG"]].sum(axis=1)
    pivot["MVA_Inngående"] = pivot[[c for c in pivot.columns if c == "TaxAmount_ING"]].sum(axis=1)

    # Netto: bruk signert-kolonner der det finnes; fallback til UTG - ING
    if "VAT_Signed_UTG" in pivot.columns or "VAT_Signed_RC_BEREGNET_UTG" in pivot.columns or "VAT_Signed_ING" in pivot.columns:
        vat_signed_out = pivot.get("VAT_Signed_UTG", 0.0) + pivot.get("VAT_Signed_RC_BEREGNET_UTG", 0.0)
        vat_signed_in = pivot.get("VAT_Signed_ING", 0.0)
        pivot["MVA_Netto_betalbar"] = (vat_signed_out + vat_signed_in).round(2)  # vat_signed_in er negativ
    else:
        pivot["MVA_Netto_betalbar"] = (pivot["MVA_Utgående"] - pivot["MVA_Inngående"]).round(2)

    # Rounding for display
    for c in pivot.columns:
        if c.startswith(("TaxBase_", "TaxAmount_", "MVA_")):
            pivot[c] = pivot[c].round(2)

    # sorter
    return pivot.sort_values(["Year", "Termin"])

# --------------------------- GUI (Streamlit) --------------------------------

st.set_page_config(page_title="MVA per termin fra SAF-T", layout="wide")
st.title("MVA per termin (fra SAF‑T) – sammenligning og analyse")

with st.sidebar:
    st.header("1) Last opp SAF‑T")
    saft_file = st.file_uploader("SAF‑T Regnskap (XML)", type=["xml"])

    st.header("2) Mapping (valgfritt)")
    st.markdown(
        "Last opp egen mapping for mva‑koder (CSV) om du vil overstyre standardene. "
        "Minst kolonnene **TaxCode** og **Direction** må være med. Valgfritt **Label**."
    )
    mapping_file = st.file_uploader("Egendefinert mapping (CSV)", type=["csv"])

    st.download_button(
        "Last ned mal for mapping (CSV)",
        data=DEFAULT_TAXCODE_MAP.to_csv(index=False).encode("utf-8"),
        file_name="mva_mapping_mal.csv",
        mime="text/csv",
    )

    st.header("3) Faktisk innrapportert (valgfritt)")
    st.markdown(
        "Last opp CSV fra Skatteetaten/Altinn eller egen konsolidering for sammenligning. "
        "Forventede kolonner: **Year, Termin, TaxCode, ReportedTaxBase, ReportedVAT**. "
        "Ekstra kolonner beholdes som metadata."
    )
    reported_file = st.file_uploader("Innrapportert (CSV)", type=["csv"])

    st.header("4) År/termin")
    selected_periodicity = st.selectbox("Periodisitet", ["6-terminer", "12-måneder", "4-kvartaler", "1-år"], index=0)

# Når fil er lastet: parse og vis
if saft_file is None:
    st.info("Last opp en SAF‑T XML‑fil for å komme i gang.")
    st.stop()

# Parse SAF-T
with st.spinner("Leser SAF‑T og bygger datasett …"):
    df_raw = parse_saft_xml(saft_file)

if df_raw.empty:
    st.error("Fant ingen linjer med TaxInformation(VAT) i filen. "
             "Sjekk at du har valgt riktig SAF‑T eksport og at filen inneholder hovedbokslinjer med mva.")
    st.stop()

# Tilpass periodisitet ved behov
if selected_periodicity != "6-terminer":
    # Regn om Termin/Label etter brukerens valg
    tmp = df_raw.apply(
        lambda r: pd.Series(month_to_termin(int(r["Month"]), periodicity=selected_periodicity),
                            index=["Termin", "TerminLabel"]),
        axis=1
    )
    df_raw["Termin"] = tmp["Termin"]
    df_raw["TerminLabel"] = tmp["TerminLabel"]

# Velg år
years = sorted(df_raw["Year"].unique().tolist())
sel_years = st.multiselect("Velg år", options=years, default=years[-1:])

if not sel_years:
    st.warning("Velg minst ett år.")
    st.stop()

# Mapping
mapping_df = load_mapping_df(mapping_file)

# Aggregater
df_per_code = summarize_by_code(df_raw, mapping_df, years=sel_years)
df_per_term = summarize_by_termin(df_per_code)

# ------------------------- Visning og nedlasting ----------------------------

tab1, tab2, tab3 = st.tabs(["Per termin (oppsummering)", "Per mva‑kode (detalj)", "Sammenligning"])

with tab1:
    st.subheader("Oppsummering per termin")
    st.dataframe(df_per_term, use_container_width=True)

    st.download_button(
        "Last ned oppsummering (CSV)",
        data=df_per_term.to_csv(index=False).encode("utf-8"),
        file_name="mva_per_termin.csv",
        mime="text/csv",
    )

with tab2:
    st.subheader("Detaljer per SAF‑T mva‑kode og termin")
    st.dataframe(df_per_code, use_container_width=True)

    st.download_button(
        "Last ned detaljer (CSV)",
        data=df_per_code.to_csv(index=False).encode("utf-8"),
        file_name="mva_per_termin_per_kode.csv",
        mime="text/csv",
    )

with tab3:
    st.subheader("Sammenligning mot innrapportert (valgfritt)")
    if reported_file is None:
        st.info("Last opp en CSV under «Faktisk innrapportert» i sidepanelet.")
    else:
        try:
            rep = pd.read_csv(reported_file)
            # normaliser typer
            rep["Year"] = rep["Year"].astype(int)
            rep["Termin"] = rep["Termin"].astype(int)
            rep["TaxCode"] = rep["TaxCode"].astype(str).str.strip().str.lstrip("0")

            # velg felter for merge
            left = df_per_code.rename(columns={"TaxBase": "CalcTaxBase", "TaxAmount": "CalcVAT"})
            merge_cols = ["Year", "Termin", "TaxCode"]

            cmp_df = rep.merge(
                left[merge_cols + ["Direction", "Label", "CalcTaxBase", "CalcVAT"]],
                how="outer",
                on=merge_cols
            )

            for c in ("ReportedTaxBase", "ReportedVAT", "CalcTaxBase", "CalcVAT"):
                if c not in cmp_df.columns:
                    cmp_df[c] = 0.0
                cmp_df[c] = cmp_df[c].fillna(0.0).astype(float).round(2)

            cmp_df["Delta_Base"] = (cmp_df["CalcTaxBase"] - cmp_df["ReportedTaxBase"]).round(2)
            cmp_df["Delta_VAT"] = (cmp_df["CalcVAT"] - cmp_df["ReportedVAT"]).round(2)

            st.dataframe(cmp_df.sort_values(["Year", "Termin", "TaxCode"]), use_container_width=True)

            st.download_button(
                "Last ned sammenligning (CSV)",
                data=cmp_df.to_csv(index=False).encode("utf-8"),
                file_name="mva_sammenligning.csv",
                mime="text/csv",
            )

            # Enkel termin-oversikt på netto (kalkulert vs rapportert)
            st.markdown("#### Termin–netto (beregnet)")
            st.dataframe(
                df_per_term[["Year", "Termin", "TerminLabel", "MVA_Utgående", "MVA_Inngående", "MVA_Netto_betalbar"]],
                use_container_width=True
            )

        except Exception as e:
            st.error(f"Feil ved lesing/merge av innrapportert CSV: {e}")

# ------------------------ Hjelpetekst i bunnen ------------------------------

with st.expander("Om data, koder og terminlogikk"):
    st.markdown(
        """
- **Termin-oppsett (default 6 terminer):** T1=jan–feb, T2=mar–apr, T3=mai–jun, T4=jul–aug, T5=sep–okt, T6=nov–des. 
  Du kan endre til måned/kvartal/år i sidepanelet.  
  Kilde for standard terminer og frister (opplysningsside): Altinn/Skatteetaten. :contentReference[oaicite:3]{index=3}

- **SAF‑T mva‑koder:** Appen bruker en start‑mapping basert på offentlig kjente beskrivelser av SAF‑T‑kodene 
  (se oversikt hos Regnskap Norge). Du kan overstyre i GUI‑et ved å laste opp egen CSV. :contentReference[oaicite:4]{index=4}

- **Koder som ikke skal rapporteres:** 0, 7, 20, 21, 22 filtreres ut i appen. Se Skatteetatens informasjonsmodell for mva‑meldingen. :contentReference[oaicite:5]{index=5}

- **Signering (netto):** I sammendraget regnes **utgående mva** positivt og **inngående mva** negativt. 
  For omvendt avgiftsplikt (fjernleverbare tjenester) vises beløpet som utgående beregnet mva. 
  Bruk egendefinert mapping hvis din løsning krever annen signering.

- **Datavariasjon:** SAF‑T fra ulike systemer kan navngi feltene under `<TaxInformation>` litt forskjellig. 
  Parseren leter etter `TaxCode`, `TaxAmount` og (hvis tilgjengelig) `TaxBase` under 
  typiske feltnavn. Om grunnlag mangler i filen, blir det satt til 0 i aggregeringen.
        """
    )

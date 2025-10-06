from __future__ import annotations

import io
import os
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from typing import List, Tuple, Optional

import fitz  # PyMuPDF
import pdfplumber

# Opsjonelle:
try:
    import camelot  # type: ignore
except Exception:
    camelot = None  # type: ignore

try:
    import pytesseract  # type: ignore
except Exception:
    pytesseract = None  # type: ignore

try:
    from pdf2image import convert_from_path  # type: ignore
except Exception:
    convert_from_path = None  # type: ignore


@dataclass
class ExtractedText:
    text: str
    blocks: List[Tuple[float, float, float, float, str]]  # (x0, y0, x1, y1, text)
    ocr_used: bool


def _get_text_blocks(path: str):
    doc = fitz.open(path)
    all_text = []
    blocks = []
    for page in doc:
        all_text.append(page.get_text() or "")
        for b in page.get_text("blocks"):
            if len(b) >= 5 and isinstance(b[4], str):
                blocks.append((b[0], b[1], b[2], b[3], b[4]))
    doc.close()
    return "\n".join(all_text), blocks


def ocr_pdf_with_cli(src_pdf: str) -> Optional[str]:
    """ OCR via ocrmypdf – nå med norsk+engelsk språk for bedre resultat. """
    if not shutil.which("ocrmypdf"):
        return None
    out_pdf = tempfile.mktemp(suffix=".pdf")
    try:
        subprocess.run(
            ["ocrmypdf", "--deskew", "--skip-text", "--language", "nor+eng", "--quiet", src_pdf, out_pdf],
            check=True
        )
        return out_pdf
    except Exception:
        return None


def ocr_pdf_with_tesseract(src_pdf: str) -> Optional[str]:
    if pytesseract is None or convert_from_path is None:
        return None
    out_pdf = tempfile.mktemp(suffix=".pdf")
    images = convert_from_path(src_pdf, dpi=300)
    doc = fitz.open()
    for img in images:
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        page = doc.new_page(width=img.width, height=img.height)
        rect = fitz.Rect(0, 0, img.width, img.height)
        page.insert_image(rect, stream=buf.getvalue())
    doc.save(out_pdf)
    doc.close()
    return out_pdf


def extract_text_blocks_from_pdf(pdf_path: str) -> ExtractedText:
    text, blocks = _get_text_blocks(pdf_path)
    if len(text.strip()) > 30:
        return ExtractedText(text=text, blocks=blocks, ocr_used=False)

    # OCR via ocrmypdf (nor+eng)
    ocr_pdf = ocr_pdf_with_cli(pdf_path)
    if ocr_pdf:
        text2, blocks2 = _get_text_blocks(ocr_pdf)
        if len(text2.strip()) > 30:
            return ExtractedText(text=text2, blocks=blocks2, ocr_used=True)

    # Fallback via pytesseract
    ocr_pdf2 = ocr_pdf_with_tesseract(pdf_path)
    if ocr_pdf2 and convert_from_path and pytesseract:
        try:
            images = convert_from_path(pdf_path, dpi=300)
            ocr_text = []
            for img in images:
                try:
                    txt = pytesseract.image_to_string(img, lang="nor+eng")
                except Exception:
                    txt = ""
                ocr_text.append(txt)
            text3 = "\n".join(ocr_text)
            return ExtractedText(text=text3, blocks=[], ocr_used=True)
        except Exception:
            pass

    return ExtractedText(text=text, blocks=blocks, ocr_used=False)


def extract_text_from_image(image_path: str) -> ExtractedText:
    if pytesseract is None:
        raise RuntimeError("pytesseract er ikke installert. Kan ikke OCR-lese bilde.")
    import PIL.Image as Image  # lazy import
    img = Image.open(image_path)
    txt = pytesseract.image_to_string(img, lang="nor+eng")
    return ExtractedText(text=txt, blocks=[], ocr_used=True)


def extract_tables(pdf_path: str):
    """Returner liste av pandas.DataFrame med tabeller."""
    import pandas as pd
    tables = []

    if camelot is not None:
        for flavor in ("lattice", "stream"):
            try:
                t = camelot.read_pdf(pdf_path, flavor=flavor, pages="1-end")
                for tbl in t:
                    df = tbl.df
                    if df is not None and df.shape[0] > 1 and df.shape[1] > 1:
                        tables.append(df)
            except Exception:
                continue

    # Fallback: pdfplumber-heuristikk
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables() or []:
                    df = pd.DataFrame(table)
                    if df.shape[0] > 1 and df.shape[1] > 1:
                        tables.append(df)
    except Exception:
        pass

    return tables

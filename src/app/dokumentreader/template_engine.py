# -*- coding: utf-8 -*-
from __future__ import annotations

# --- bootstrap ---
import os, sys, re, json
from dataclasses import dataclass, asdict
from typing import List, Dict, Optional, Tuple, Any

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../"))
if BASE not in sys.path:
    sys.path.insert(0, BASE)
# -----------------

import fitz  # PyMuPDF
try:
    import yaml  # pyyaml
except Exception as _e:
    yaml = None  # type: ignore

TEMPLATES_ROOT = os.path.join(BASE, "app", "dokumentreader", "templates")
os.makedirs(TEMPLATES_ROOT, exist_ok=True)

# ----------------- datamodell -----------------

@dataclass
class Rule:
    field: str
    anchor_text: str
    page: Optional[int]          # 0-basert side (None = valgfri)
    search: str                  # 'right' | 'below'
    max_dx: float
    max_dy: float
    band: float                  # “samme linje/kolonne”-toleranse
    anchor_hint_y: Optional[float] = None  # hjelper å velge riktig forekomst
    value_regex: Optional[str] = None      # f.eks. (?i)ja|nei / tall
    value_hint: Optional[str] = None       # eksempelverdi (visning)

@dataclass
class Template:
    profile: str                # 'invoice' | 'financials_no' | 'tax_return' | ...
    name: str
    match_any: List[str]        # enkle “match-ord” for gjenkjenning
    rules: List[Rule]

# ----------------- utilities -----------------

def _load_yaml(path: str) -> dict:
    if yaml is None:
        raise RuntimeError("PyYAML er ikke installert. installer 'pyyaml'.")
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}

def _dump_yaml(path: str, data: dict) -> None:
    if yaml is None:
        raise RuntimeError("PyYAML er ikke installert. installer 'pyyaml'.")
    with open(path, "w", encoding="utf-8") as f:
        yaml.safe_dump(data, f, sort_keys=False, allow_unicode=True)

def _to_lines(words: List[Tuple[float,float,float,float,str]]) -> List[Tuple[fitz.Rect,str]]:
    """
    Gruppér ord til linjer. Bruk y0-gitter for å samle ord som står på samme rad.
    """
    if not words:
        return []
    groups: Dict[int, List[Tuple[float,float,float,float,str]]] = {}
    for w in words:
        x0,y0,x1,y1,tx = w[:5]
        key = int(round(y0/2.0))  # 2 px gitter – justeres ved behov
        groups.setdefault(key, []).append((x0,y0,x1,y1,tx))
    out: List[Tuple[fitz.Rect,str]] = []
    for g in groups.values():
        g.sort(key=lambda r: (r[1], r[0]))
        x0=min(v[0] for v in g); y0=min(v[1] for v in g)
        x1=max(v[2] for v in g); y1=max(v[3] for v in g)
        text = " ".join(v[4] for v in g)
        out.append((fitz.Rect(x0,y0,x1,y1), text))
    out.sort(key=lambda r: (r[0].y0, r[0].x0))
    return out

def _line_at_click(page: fitz.Page, x: float, y: float) -> Tuple[fitz.Rect, str]:
    """
    Finn *linjen* (ikke blokken) under klikkpunktet (x,y) i PDF-koordinater.
    Fallback: nærmeste linje i |y|.
    """
    words = page.get_text("words") or []
    lines = _to_lines([(w[0],w[1],w[2],w[3],w[4]) for w in words if len(w)>=5])
    for r, tx in lines:
        if r.contains(fitz.Point(x, y)):
            return r, tx
    # nærmeste i vertikal retning
    if lines:
        r, tx = min(lines, key=lambda lt: abs((lt[0].y0+lt[0].y1)/2 - y))
        return r, tx
    # fallback: hele blokken
    blks = page.get_text("blocks") or []
    best = None
    for b in blks:
        if len(b)>=5 and isinstance(b[4], str):
            R = fitz.Rect(b[0],b[1],b[2],b[3])
            if best is None or (abs((R.y0+R.y1)/2 - y) + abs((R.x0+R.x1)/2 - x)) < \
                               (abs((best[0].y0+best[0].y1)/2 - y) + abs((best[0].x0+best[0].x1)/2 - x)):
                best = (R, b[4])
    if best:
        return best
    return fitz.Rect(x-2,y-2,x+2,y+2), ""

def _zone_right(anchor: fitz.Rect, max_dx: float, band: float) -> fitz.Rect:
    cy = (anchor.y0 + anchor.y1)/2
    return fitz.Rect(anchor.x1, cy-band/2, anchor.x1+max_dx, cy+band/2)

def _zone_below(anchor: fitz.Rect, max_dy: float, band: float) -> fitz.Rect:
    cx = (anchor.x0 + anchor.x1)/2
    return fitz.Rect(cx-band/2, anchor.y1, cx+band/2, anchor.y1+max_dy)

def _pick_line_in_zone(page: fitz.Page, zone: fitz.Rect, prefer_y: Optional[float], prefer_x: Optional[float]) -> Optional[Tuple[fitz.Rect,str]]:
    words = page.get_text("words") or []
    lines = _to_lines([(w[0],w[1],w[2],w[3],w[4]) for w in words])
    cands = [(r,tx) for (r,tx) in lines if r.intersects(zone)]
    if not cands:
        return None
    if prefer_y is not None:
        return min(cands, key=lambda lt: abs((lt[0].y0+lt[0].y1)/2 - prefer_y))
    if prefer_x is not None:
        return min(cands, key=lambda lt: abs((lt[0].x0+lt[0].x1)/2 - prefer_x))
    return cands[0]

def _norm_text(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def _default_value_regex(sample: str) -> Optional[str]:
    t = (sample or "").strip()
    if not t:
        return None
    if re.fullmatch(r"(?i)ja|nei", t) or t.lower() in {"ja","nei"}:
        return r"(?i)ja|nei"
    if re.fullmatch(r"[0-9\s.,\u00A0\u202F\-]+", t):
        return r"[0-9\s.,\u00A0\u202F\-]+"
    return None

# ----------------- publikt API -----------------

def build_rule_from_clicks(doc: fitz.Document, page_idx: int,
                           anchor_xy: Tuple[float,float],
                           value_xy: Tuple[float,float],
                           field_name: str) -> Rule:
    page = doc[page_idx]
    ax, ay = anchor_xy; vx, vy = value_xy
    a_rect, a_text = _line_at_click(page, ax, ay)
    v_rect, v_text = _line_at_click(page, vx, vy)

    # retning og toleranser
    if v_rect.x0 >= a_rect.x1:
        search = "right"
        max_dx = max(60.0, v_rect.x1 - a_rect.x1 + 20.0)
        band   = max(a_rect.height, v_rect.height) * 1.6
        max_dy = 0.0
    else:
        search = "below"
        max_dy = max(40.0, v_rect.y1 - a_rect.y1 + 20.0)
        band   = max(a_rect.width, v_rect.width) * 0.8
        max_dx = 0.0

    return Rule(
        field=field_name,
        anchor_text=_norm_text(a_text),
        page=page_idx,
        search=search,
        max_dx=max_dx,
        max_dy=max_dy,
        band=band,
        anchor_hint_y=(a_rect.y0+a_rect.y1)/2,
        value_regex=_default_value_regex(_norm_text(v_text)),
        value_hint=_norm_text(v_text)
    )

def save_template(profile: str, name: str, match_any: List[str], rules: List[Rule]) -> str:
    data = {
        "profile": profile,
        "name": name,
        "match": {"any": match_any or []},
        "anchors": [asdict(r) for r in rules],
    }
    out = os.path.join(TEMPLATES_ROOT, f"{profile}__{re.sub(r'[^a-zA-Z0-9_-]+','_',name)}.yaml")
    _dump_yaml(out, data)
    return out

def load_templates(profile: Optional[str] = None) -> List[Template]:
    out: List[Template] = []
    for fn in os.listdir(TEMPLATES_ROOT):
        if not fn.lower().endswith(".yaml"):
            continue
        data = _load_yaml(os.path.join(TEMPLATES_ROOT, fn))
        prof = str(data.get("profile") or "")
        if profile and prof != profile:
            continue
        rules = [
            Rule(
                field=a.get("field",""),
                anchor_text=a.get("anchor_text",""),
                page=a.get("page"),
                search=a.get("search","right"),
                max_dx=float(a.get("max_dx", 200.0)),
                max_dy=float(a.get("max_dy", 80.0)),
                band=float(a.get("band", 24.0)),
                anchor_hint_y=a.get("anchor_hint_y"),
                value_regex=a.get("value_regex"),
                value_hint=a.get("value_hint")
            )
            for a in (data.get("anchors") or [])
        ]
        out.append(Template(
            profile=prof,
            name=str(data.get("name") or os.path.splitext(fn)[0]),
            match_any=[str(x) for x in (data.get("match",{}).get("any") or [])],
            rules=rules
        ))
    return out

def _text_contains_any(text: str, words: List[str]) -> bool:
    t = text.lower()
    return any((w or "").lower() in t for w in words)

def apply_templates(pdf_path: str, profile: str,
                    prefer_template: Optional[str] = None) -> Dict[str, str]:
    """
    Kjør maler for en PDF og returner {felt: verdi}.
    Velger første mal som matcher på 'match_any' (eller 'prefer_template' ved navn-match).
    """
    doc = fitz.open(pdf_path)
    try:
        page0_text = ""
        try:
            page0_text = doc[0].get_text() or ""
        except Exception:
            pass

        tpls = load_templates(profile)
        if prefer_template:
            tpls = [t for t in tpls if t.name == prefer_template] + [t for t in tpls if t.name != prefer_template]
        chosen: Optional[Template] = None
        for t in tpls:
            if not t.match_any or _text_contains_any(page0_text, t.match_any):
                chosen = t
                break
        if not chosen:
            return {}

        results: Dict[str, str] = {}
        for r in chosen.rules:
            pidx = r.page if (r.page is not None and 0 <= r.page < len(doc)) else 0
            page = doc[pidx]
            # finn anker-forekomster
            hits = page.search_for(r.anchor_text) or []
            if not hits:
                continue
            # velg forekomst nær 'anchor_hint_y' hvis vi har
            if r.anchor_hint_y is not None:
                rect = min(hits, key=lambda R: abs((R.y0+R.y1)/2 - r.anchor_hint_y))
            else:
                rect = hits[0]

            if r.search == "right":
                zone = _zone_right(rect, max(10.0, r.max_dx), max(8.0, r.band))
                cand = _pick_line_in_zone(page, zone, prefer_y=(rect.y0+rect.y1)/2, prefer_x=None)
            else:
                zone = _zone_below(rect, max(10.0, r.max_dy), max(8.0, r.band))
                cand = _pick_line_in_zone(page, zone, prefer_y=None, prefer_x=(rect.x0+rect.x1)/2)

            if not cand:
                continue

            _, text = cand
            text = _norm_text(text)

            if r.value_regex:
                m = re.search(r.value_regex, text)
                if m:
                    text = _norm_text(m.group(0))
                else:
                    # ingen regex-treff – la oss likevel ta hele linjen (sikkerhet)
                    pass

            results[r.field] = text

        return results
    finally:
        try: doc.close()
        except: pass

def debug_test_template(pdf_path: str, tpl: Template) -> Dict[str, Any]:
    """
    Returner både felt og debug-info (soner) for Admin/visning.
    """
    doc = fitz.open(pdf_path)
    zones = []
    values = {}
    try:
        for r in tpl.rules:
            pidx = r.page if (r.page is not None and 0 <= r.page < len(doc)) else 0
            page = doc[pidx]
            hits = page.search_for(r.anchor_text) or []
            if not hits:
                continue
            if r.anchor_hint_y is not None:
                rect = min(hits, key=lambda R: abs((R.y0+R.y1)/2 - r.anchor_hint_y))
            else:
                rect = hits[0]
            if r.search == "right":
                zone = _zone_right(rect, max(10.0, r.max_dx), max(8.0, r.band))
                cand = _pick_line_in_zone(page, zone, (rect.y0+rect.y1)/2, None)
            else:
                zone = _zone_below(rect, max(10.0, r.max_dy), max(8.0, r.band))
                cand = _pick_line_in_zone(page, zone, None, (rect.x0+rect.x1)/2)

            zones.append({"page": pidx, "anchor": [rect.x0,rect.y0,rect.x1,rect.y1],
                          "zone": [zone.x0,zone.y0,zone.x1,zone.y1], "field": r.field})

            if cand:
                _, tx = cand
                tx = _norm_text(tx)
                if r.value_regex:
                    m = re.search(r.value_regex, tx)
                    if m:
                        tx = _norm_text(m.group(0))
                values[r.field] = tx
        return {"values": values, "zones": zones, "template": asdict(tpl)}
    finally:
        try: doc.close()
        except: pass

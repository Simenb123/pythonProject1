from __future__ import annotations
import os, tempfile, webbrowser
from typing import Dict, List, Tuple, Set

from .db import get_owners_agg_owner, get_children_agg_company
from . import settings as S

# ---- Enkle datatyper ----
NodeId = str  # enten orgnr eller "U:<navn>" for privatperson uten orgnr
Edge   = Tuple[NodeId, NodeId, str]  # (src, dst, label)

# ---- Bygg grafdatastruktur ved å slå opp i DB ----
def _gather_graph(conn, root_orgnr: str, root_name: str, mode: str, max_up: int, max_down: int
                  ) -> Tuple[Dict[NodeId, str], List[Edge]]:
    labels: Dict[NodeId, str] = {}
    edges:  List[Edge] = []

    labels[root_orgnr] = f"{root_name}\n({root_orgnr})"

    visited_up:   Set[NodeId] = set()
    visited_down: Set[NodeId] = set()

    def up(cur: str, depth: int):
        if depth > max_up or cur in visited_up: return
        visited_up.add(cur)
        for owner_orgnr, owner_name, *_rest, pct in get_owners_agg_owner(conn, cur):
            nid: NodeId = owner_orgnr or f"U:{owner_name}"
            labels.setdefault(nid, f"{owner_name}\n({owner_orgnr or '–'})")
            edges.append((nid, cur, f"{pct:.2f}%" if pct is not None else ""))
            if owner_orgnr: up(owner_orgnr, depth + 1)

    def down(cur: str, depth: int):
        if depth > max_down or cur in visited_down: return
        visited_down.add(cur)
        for child_orgnr, child_name, *_rest, pct in get_children_agg_company(conn, cur):
            if not child_orgnr: continue
            labels.setdefault(child_orgnr, f"{child_name}\n({child_orgnr})")
            edges.append((cur, child_orgnr, f"{pct:.2f}%" if pct is not None else ""))
            down(child_orgnr, depth + 1)

    if mode in ("both","up"):   up(root_orgnr,   1)
    if mode in ("both","down"): down(root_orgnr, 1)
    return labels, edges

# ---- Enkel layouter (lagvis i Y, jevn fordeling i X) ----
def _layout(labels: Dict[NodeId, str], edges: List[Edge], root: NodeId,
            max_up: int, max_down: int) -> Dict[NodeId, Tuple[int,int]]:
    # Lagdybder: negative = oppstrøms, 0 = root, positive = nedstrøms
    layer: Dict[NodeId, int] = {root: 0}
    # opp
    frontier = [root]
    for d in range(1, max_up+1):
        nxt = set()
        for _, dst, _ in edges:
            if dst in frontier:
                # finn kilder (eier -> dst)
                for src, dst2, _ in edges:
                    if dst2 == dst:
                        if src not in layer:
                            layer[src] = -d
                            nxt.add(src)
        frontier = list(nxt)
    # ned
    frontier = [root]
    for d in range(1, max_down+1):
        nxt = set()
        for src, _, _ in edges:
            if src in frontier:
                for src2, dst, _ in edges:
                    if src2 == src:
                        if dst not in layer:
                            layer[dst] = d
                            nxt.add(dst)
        frontier = list(nxt)

    # grupper per lag
    per_layer: Dict[int, List[NodeId]] = {}
    for nid, d in layer.items():
        per_layer.setdefault(d, []).append(nid)
    for d in per_layer:
        per_layer[d].sort(key=lambda n: labels[n])

    # posisjoner
    pos: Dict[NodeId, Tuple[int,int]] = {}
    W  = 180  # nodebredde (estimat)
    H  = 60   # nodehøyde
    XS = 220  # X-spacing
    YS = 120  # Y-spacing

    # root midt
    pos[root] = (0, 0)

    # opp (negative lag)
    for depth in sorted([d for d in per_layer if d < 0]):
        nodes = per_layer[depth]
        n = len(nodes)
        for i, nid in enumerate(nodes):
            x = (i - (n-1)/2) * XS
            y = depth * YS
            pos[nid] = (int(x), int(y))

    # ned (positive lag)
    for depth in sorted([d for d in per_layer if d > 0]):
        nodes = per_layer[depth]
        n = len(nodes)
        for i, nid in enumerate(nodes):
            x = (i - (n-1)/2) * XS
            y = depth * YS
            pos[nid] = (int(x), int(y))

    return pos

# ---- HTML+SVG generator ----
def _svg_html(labels: Dict[NodeId, str], edges: List[Edge], pos: Dict[NodeId, Tuple[int,int]],
              root: NodeId, title: str) -> str:
    # beregn viewbox
    xs = [x for x, _ in pos.values()] + [0]
    ys = [y for _, y in pos.values()] + [0]
    # Beregn symmetrisk bredde og høyde slik at rot (0,0) ligger midt i SVG.
    max_abs_x = max(abs(x) for x in xs)
    max_abs_y = max(abs(y) for y in ys)
    margin_x = 200  # ekstra luft på sidene
    margin_y = 120  # ekstra luft oppe og nede
    width  = max(900, int(2 * (max_abs_x + margin_x)))
    height = max(600, int(2 * (max_abs_y + margin_y)))

    def esc(s: str) -> str:
        return (s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"))

    # Bygg noder
    node_svgs: List[str] = []
    for nid, label in labels.items():
        x, y = pos[nid]
        # Plasser noder midtstilt: legg til halv bredde/høyde på koordinatene (rot ligger på 0,0)
        X = int(x + width // 2)
        Y = int(y + height // 2)
        is_company = (nid == root) or nid.isdigit()
        fill = "#e9ecef" if is_company else "#ffffff"
        stroke = "#495057" if is_company else "#6c757d"
        node_svgs.append(
            f'<g class="node" transform="translate({X-80},{Y-30})">'
            f'<rect x="0" y="0" rx="8" ry="8" width="160" height="60" '
            f'style="fill:{fill};stroke:{stroke};stroke-width:1.2"/>'
            f'<text x="80" y="25" text-anchor="middle" font-family="Arial, Helvetica, sans-serif" '
            f'font-size="12" fill="#212529">{esc(label).splitlines()[0]}</text>'
            f'<text x="80" y="43" text-anchor="middle" font-family="Arial, Helvetica, sans-serif" '
            f'font-size="11" fill="#6c757d">{esc(label).splitlines()[1] if "\\n" in label else ""}</text>'
            f"</g>"
        )

    # Bygg kanter
    edge_svgs: List[str] = []
    for src, dst, lbl in edges:
        # hent posisjon (x,y) for kilder og destinasjoner, og midtstill dem i SVG
        x1, y1 = pos[src]; X1 = int(x1 + width // 2); Y1 = int(y1 + height // 2)
        x2, y2 = pos[dst]; X2 = int(x2 + width // 2); Y2 = int(y2 + height // 2)
        # piler går fra nederst på src (Y + 30) til øverst på dst (Y - 30)
        x1a, y1a = X1, Y1 + 30
        x2a, y2a = X2, Y2 - 30
        edge_svgs.append(
            f'<path d="M{x1a},{y1a} C{x1a},{(y1a+y2a)//2} {x2a},{(y1a+y2a)//2} {x2a},{y2a}" '
            f'style="fill:none;stroke:#6c757d;stroke-width:1.2" marker-end="url(#arrow)"/>'
            f'<text x="{(x1a+x2a)//2}" y="{(y1a+y2a)//2 - 4}" text-anchor="middle" '
            f'font-size="11" font-family="Arial, Helvetica, sans-serif" fill="#495057">{esc(lbl)}</text>'
        )

    html = f"""<!doctype html>
<html lang="no">
<meta charset="utf-8"/>
<title>{esc(title)}</title>
<style>
/* La siden kunne scrolles dersom SVG er større enn viewport */
body{{margin:0;background:#f8f9fa; overflow:auto}}
</style>
<!-- Sett eksplisitt bredde/høyde på SVG slik at hele grafen kan rulles -->
<svg width="{width}" height="{height}" viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg">
<defs>
  <marker id="arrow" markerWidth="10" markerHeight="7" refX="10" refY="3.5" orient="auto">
    <polygon points="0 0, 10 3.5, 0 7" style="fill:#6c757d;"/>
  </marker>
</defs>
<rect x="0" y="0" width="{width}" height="{height}" fill="#f8f9fa"/>
<text x="{width//2}" y="24" text-anchor="middle" font-size="16"
      font-family="Arial, Helvetica, sans-serif" fill="#212529">{esc(title)}</text>
{''.join(edge_svgs)}
{''.join(node_svgs)}
</svg>
</html>"""
    return html

# ---- Offentlig API ----
def render_graph(conn,
                 company_orgnr: str,
                 company_name: str,
                 mode: str = "both",
                 max_up: int = S.MAX_DEPTH_UP,
                 max_down: int = S.MAX_DEPTH_DOWN) -> str | None:
    """
    Generer orgkart som enkel HTML+SVG (ingen Graphviz/PyVis).
    Returnerer sti til .html-filen og åpner den i nettleser.
    """
    labels, edges = _gather_graph(conn, company_orgnr, company_name, mode, max_up, max_down)
    title = f"Eierskapstre for {company_name} ({company_orgnr}) – {mode}"
    pos   = _layout(labels, edges, company_orgnr, max_up, max_down)
    html  = _svg_html(labels, edges, pos, company_orgnr, title)
    outdir = tempfile.gettempdir()
    path   = os.path.join(outdir, f"eierskap_{company_orgnr}.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    try:
        webbrowser.open_new_tab(path)
    except Exception:
        pass
    return path

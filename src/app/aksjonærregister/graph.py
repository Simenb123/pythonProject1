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

# ---- Hierarkisk layouter med rekursiv plassering av foreldre og barn ----
def _layout(labels: Dict[NodeId, str], edges: List[Edge], root: NodeId,
            max_up: int, max_down: int) -> Dict[NodeId, Tuple[int, int]]:
    """
    Beregn posisjoner for noder i et todimensjonalt orgkart. Vi bruker en
    lagvis tilnærming: nodene grupperes etter avstand (lag) fra rotnoden,
    og innen hvert lag sorteres de ved hjelp av barysenterheuristikk for å
    redusere kryssende kanter. Deretter plasseres de jevnt langs x-aksen.

    Roten ligger på (0,0). Lag > 0 er nedstrøms (barn), lag < 0 er oppstrøms
    (foreldre). Y-koordinaten er proporsjonal med laget. X-koordinaten settes
    slik at laget sentreres rundt rotaksen og grupper med samme foreldre
    havner nær hverandre.
    """
    # Beregn lag for hver node: negativ for oppstrøms, 0 for root, positiv for nedstrøms
    layer: Dict[NodeId, int] = {root: 0}
    # Utvid oppover
    frontier = [root]
    for d in range(1, max_up + 1):
        nxt: Set[NodeId] = set()
        for src, dst, _ in edges:
            if dst in frontier and src not in layer:
                layer[src] = -d
                nxt.add(src)
        frontier = list(nxt)
    # Utvid nedover
    frontier = [root]
    for d in range(1, max_down + 1):
        nxt: Set[NodeId] = set()
        for src, dst, _ in edges:
            if src in frontier and dst not in layer:
                layer[dst] = d
                nxt.add(dst)
        frontier = list(nxt)

    # Bygg per-lag-lister og initialiser sortering alfabetisk
    layers: Dict[int, List[NodeId]] = {}
    for nid, depth in layer.items():
        layers.setdefault(depth, []).append(nid)
    for depth in layers:
        layers[depth].sort(key=lambda n: labels.get(n, ""))

    # Bygg adjacens-lister for nedstrøms pass: foreldre til barn (lag >=0)
    parents_down: Dict[NodeId, List[NodeId]] = {}
    for src, dst, _ in edges:
        ds = layer.get(src, 0)
        dd = layer.get(dst, 0)
        # Nedstrøms kant
        if ds >= 0 and dd > ds:
            parents_down.setdefault(dst, []).append(src)
    # Sorter foreldre alfabetisk for konsistens
    for node, par in parents_down.items():
        par.sort(key=lambda n: labels.get(n, ""))

    # Barycenter-sortering for nedstrøms nivåer
    for k in range(1, max_down + 1):
        nodes = layers.get(k, [])
        if not nodes:
            continue
        prev_layer_nodes = layers.get(k - 1, [root])
        # Lag en index-oppslag for foreldre i forrige lag
        index_prev: Dict[NodeId, int] = {nid: i for i, nid in enumerate(prev_layer_nodes)}
        # Beregn barysenter for hvert node basert på foreldre
        bary: Dict[NodeId, float] = {}
        for node in nodes:
            par = parents_down.get(node, [])
            # Dersom flere foreldre, ta gjennomsnitt av indeksene; ellers bruk egen indeks
            if par:
                positions = [index_prev[p] for p in par if p in index_prev]
                if positions:
                    bary[node] = sum(positions) / len(positions)
                else:
                    bary[node] = nodes.index(node)
            else:
                bary[node] = nodes.index(node)
        # Sorter nodene etter barysenter (stiger)
        nodes.sort(key=lambda n: bary[n])
        layers[k] = nodes

    # Bygg adjacens-lister for oppstrøms pass: foreldre->barn i oppstrømsdelen
    children_up: Dict[NodeId, List[NodeId]] = {}
    for src, dst, _ in edges:
        ds = layer.get(src, 0)
        dd = layer.get(dst, 0)
        # Oppstrøms kant: kilde ligger høyere (lavere layer-verdi) enn dest
        if ds < 0 and dd <= 0:
            children_up.setdefault(src, []).append(dst)
    # Sorter barna alfabetisk for konsistens
    for node, kids in children_up.items():
        kids.sort(key=lambda n: labels.get(n, ""))

    # Bygg lag for oppstrøms (positiv indeks for enkelhet)
    layers_up: Dict[int, List[NodeId]] = {0: [root]}
    for nid, depth in layer.items():
        if depth < 0:
            layers_up.setdefault(-depth, []).append(nid)
    for depth in layers_up:
        layers_up[depth].sort(key=lambda n: labels.get(n, ""))

    # Barycenter-sortering for oppstrøms lag
    for k in range(1, max_up + 1):
        nodes = layers_up.get(k, [])
        if not nodes:
            continue
        prev_layer_nodes = layers_up.get(k - 1, [root])
        index_prev: Dict[NodeId, int] = {nid: i for i, nid in enumerate(prev_layer_nodes)}
        bary: Dict[NodeId, float] = {}
        for node in nodes:
            kids = children_up.get(node, [])
            if kids:
                positions = [index_prev[c] for c in kids if c in index_prev]
                if positions:
                    bary[node] = sum(positions) / len(positions)
                else:
                    bary[node] = nodes.index(node)
            else:
                bary[node] = nodes.index(node)
        nodes.sort(key=lambda n: bary[n])
        layers_up[k] = nodes

    # Tildel koordinater: x-koordinater er evenly spaced per lag
    pos: Dict[NodeId, Tuple[int, int]] = {root: (0, 0)}
    XS = 220
    YS = 120
    # Nedstrøms lag (positive)
    for depth in range(1, max_down + 1):
        nodes = layers.get(depth, [])
        n = len(nodes)
        if n == 0:
            continue
        x_start = -((n - 1) / 2.0) * XS
        for i, nid in enumerate(nodes):
            x = x_start + i * XS
            y = depth * YS
            pos[nid] = (int(round(x)), int(round(y)))
    # Oppstrøms lag (negative)
    for k in range(1, max_up + 1):
        nodes = layers_up.get(k, [])
        n = len(nodes)
        if n == 0:
            continue
        x_start = -((n - 1) / 2.0) * XS
        for i, nid in enumerate(nodes):
            x = x_start + i * XS
            y = -k * YS
            pos[nid] = (int(round(x)), int(round(y)))

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
        # Split label into up to two lines to avoid backslash inside f-string expression.
        _escaped_label = esc(label)
        _label_lines = _escaped_label.splitlines()
        if len(_label_lines) < 2:
            _label_lines.append("")
        node_svgs.append(
            f'<g class="node" data-id="{nid}" transform="translate({X-80},{Y-30})">'
            f'<rect x="0" y="0" rx="8" ry="8" width="160" height="60" '
            f'style="fill:{fill};stroke:{stroke};stroke-width:1.2"/>'
            f'<text x="80" y="25" text-anchor="middle" font-family="Arial, Helvetica, sans-serif" '
            f'font-size="12" fill="#212529">{_label_lines[0]}</text>'
            f'<text x="80" y="43" text-anchor="middle" font-family="Arial, Helvetica, sans-serif" '
            f'font-size="11" fill="#6c757d">{_label_lines[1]}</text>'
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
            f'<path data-src="{src}" data-dst="{dst}" '
            f'd="M{x1a},{y1a} C{x1a},{(y1a+y2a)//2} {x2a},{(y1a+y2a)//2} {x2a},{y2a}" '
            f'style="fill:none;stroke:#6c757d;stroke-width:1.2" marker-end="url(#arrow)"/>'
            f'<text data-src="{src}" data-dst="{dst}" x="{(x1a+x2a)//2}" y="{(y1a+y2a)//2 - 4}" text-anchor="middle" '
            f'font-size="11" font-family="Arial, Helvetica, sans-serif" fill="#495057">{esc(lbl)}</text>'
        )

    # Interaktiv skript for å flytte noder med mus og oppdatere kanter. Dette legges til på slutten av HTMLen.
    script = """
<script>
(function() {
  const nodeWidth = 160;
  const nodeHeight = 60;
  function parseTransform(transform) {
    const match = /translate\(([-0-9.]+),([-0-9.]+)\)/.exec(transform);
    return {x: parseFloat(match[1]), y: parseFloat(match[2])};
  }
  const nodeElems = {};
  document.querySelectorAll('g.node').forEach(node => {
    const id = node.getAttribute('data-id');
    nodeElems[id] = node;
  });
  const edgesList = [];
  document.querySelectorAll('path[data-src]').forEach(pathElem => {
    const src = pathElem.getAttribute('data-src');
    const dst = pathElem.getAttribute('data-dst');
    const textElem = document.querySelector(`text[data-src="${src}"][data-dst="${dst}"]`);
    edgesList.push({path: pathElem, text: textElem, src: src, dst: dst});
  });
  let draggingNode = null;
  const dragStart = {x: 0, y: 0};
  const nodeStart = {x: 0, y: 0};
  function updateEdges(id) {
    edgesList.forEach(edge => {
      if (edge.src === id || edge.dst === id) {
        const srcNode = nodeElems[edge.src];
        const dstNode = nodeElems[edge.dst];
        const tSrc = parseTransform(srcNode.getAttribute('transform'));
        const tDst = parseTransform(dstNode.getAttribute('transform'));
        const x1 = tSrc.x + nodeWidth / 2;
        const y1 = tSrc.y + nodeHeight;
        const x2 = tDst.x + nodeWidth / 2;
        const y2 = tDst.y;
        const midY = (y1 + y2) / 2;
        edge.path.setAttribute('d', `M${x1},${y1} C${x1},${midY} ${x2},${midY} ${x2},${y2}`);
        edge.text.setAttribute('x', (x1 + x2) / 2);
        edge.text.setAttribute('y', midY - 4);
      }
    });
  }
  document.querySelectorAll('g.node').forEach(node => {
    node.addEventListener('mousedown', function(evt) {
      draggingNode = this;
      const trans = parseTransform(this.getAttribute('transform'));
      nodeStart.x = trans.x;
      nodeStart.y = trans.y;
      dragStart.x = evt.clientX;
      dragStart.y = evt.clientY;
      evt.preventDefault();
    });
  });
  document.addEventListener('mousemove', function(evt) {
    if (!draggingNode) return;
    const dx = evt.clientX - dragStart.x;
    const dy = evt.clientY - dragStart.y;
    const newX = nodeStart.x + dx;
    const newY = nodeStart.y + dy;
    draggingNode.setAttribute('transform', `translate(${newX},${newY})`);
    const id = draggingNode.getAttribute('data-id');
    updateEdges(id);
  });
  document.addEventListener('mouseup', function(evt) {
    if (draggingNode) {
      draggingNode = null;
    }
  });
})();
</script>
"""

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
{script}
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

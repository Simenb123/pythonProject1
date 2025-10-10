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
    # Avstand mellom noder (horisontalt) og mellom nivåer (vertikalt). Øk disse
    # for å få mer luft i orgkartet. Standardverdiene 220 og 120 gir en tett
    # layout; vi øker dem her for bedre oversikt.
    XS = 280
    YS = 160
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
    # Øk marginene rundt grafen for å gi mer luft når noder og lag spres
    margin_x = 250  # ekstra luft på sidene
    margin_y = 200  # ekstra luft oppe og nede
    width  = max(900, int(2 * (max_abs_x + margin_x)))
    height = max(600, int(2 * (max_abs_y + margin_y)))

    def esc(s: str) -> str:
        return (s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"))

    # Bygg noder med fargekoder og figurer avhengig av om det er selskap eller privatperson.
    node_svgs: List[str] = []
    for nid, label in labels.items():
        x, y = pos[nid]
        # Plasser noder midtstilt: legg til halv bredde/høyde på koordinatene (rot ligger på 0,0)
        X = int(x + width // 2)
        Y = int(y + height // 2)
        # Vi anser noder som selskaper dersom de er rotnoden eller består
        # av nøyaktig 9 sifre (norsk organisasjonsnummer). Andre tallstrenger
        # (f.eks. 11-sifret fødselsnummer eller 4-sifret fødselsår) behandles
        # som privatpersoner.
        is_company = (nid == root) or (nid.isdigit() and len(nid) == 9)
        # Definer farger: selskaper får nøytral farge, personer får blå tone
        if is_company:
            # Selskaper: lysegrå bakgrunn, mørkere ramme
            fill = "#f0f0f0"
            stroke = "#6c757d"
        else:
            # Privatpersoner: lys pastellblå bakgrunn, tydelig blå ramme
            fill = "#d6e4f0"
            stroke = "#4879c0"
        # Tykkere kantlinje for rotnode for å fremheve den
        border_width = 2.0 if nid == root else 1.2
        # Split label i inntil to linjer
        _escaped_label = esc(label)
        _label_lines = _escaped_label.splitlines()
        if len(_label_lines) < 2:
            _label_lines.append("")
        tooltip = _label_lines[0] + (" – " + _label_lines[1] if _label_lines[1] else "")
        # Velg form: rektangel for selskap, ellipse for privatperson
        if is_company:
            shape_tag = (
                f'<rect x="0" y="0" rx="8" ry="8" width="160" height="60" '
                f'style="fill:{fill};stroke:{stroke};stroke-width:{border_width}"/>'
            )
        else:
            # Ellipse har samme midtpunkt og størrelse som rektangelet
            shape_tag = (
                f'<ellipse cx="80" cy="30" rx="80" ry="30" '
                f'style="fill:{fill};stroke:{stroke};stroke-width:{border_width}"/>'
            )
        node_svgs.append(
            f'<g class="node" data-id="{nid}" transform="translate({X-80},{Y-30})">'
            f'<title>{tooltip}</title>'
            f'{shape_tag}'
            f'<text x="80" y="25" text-anchor="middle" font-family="Arial, Helvetica, sans-serif" '
            f'font-size="12" fill="#212529">{_label_lines[0]}</text>'
            f'<text x="80" y="43" text-anchor="middle" font-family="Arial, Helvetica, sans-serif" '
            f'font-size="11" fill="#6c757d">{_label_lines[1]}</text>'
            f"</g>"
        )

    # Bygg kanter, tilpass linjetykkelse etter eierandel (høyere prosent => tykkere linje)
    edge_svgs: List[str] = []
    for src, dst, lbl in edges:
        # hent posisjon (x,y) for kilder og destinasjoner, og midtstill dem i SVG
        x1, y1 = pos[src]; X1 = int(x1 + width // 2); Y1 = int(y1 + height // 2)
        x2, y2 = pos[dst]; X2 = int(x2 + width // 2); Y2 = int(y2 + height // 2)
        # piler går fra nederst på src (Y + 30) til øverst på dst (Y - 30)
        x1a, y1a = X1, Y1 + 30
        x2a, y2a = X2, Y2 - 30
        # Beregn linjetykkelse og farge ut fra eierandel (prosent). Dersom
        # prosent ikke finnes, behold standard verdi. Høyere eierandel gir
        # tykkere linje. Fargekoder: ≥50 % = grønn, 10–49 % = gul, <10 % = rød.
        stroke_width = 1.2
        stroke_color = "#6c757d"
        pct_val = None
        if lbl:
            s = lbl.strip().rstrip('%').replace(',', '.')
            try:
                pct_val = float(s)
            except Exception:
                pct_val = None
        if pct_val is not None:
            # Linjetykkelse: 1 px + prosent/50 → 100 % ≈ 3 px, 50 % ≈ 2 px
            stroke_width = 1.0 + pct_val / 50.0
            # Velg farge basert på intervall
            if pct_val >= 50.0:
                stroke_color = "#28a745"  # grønn
            elif pct_val >= 10.0:
                stroke_color = "#ffc107"  # gul
            else:
                stroke_color = "#dc3545"  # rød
        edge_svgs.append(
            f'<path data-src="{src}" data-dst="{dst}" '
            f'd="M{x1a},{y1a} C{x1a},{(y1a+y2a)//2} {x2a},{(y1a+y2a)//2} {x2a},{y2a}" '
            f'style="fill:none;stroke:{stroke_color};stroke-width:{stroke_width}" marker-end="url(#arrow)"/>'
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

  // Zooming and panning
  const graph = document.getElementById('graph');
  const svg = document.querySelector('svg');
  let scale = 1;
  let translateX = 0;
  let translateY = 0;
  let panning = false;
  let panStartX = 0;
  let panStartY = 0;
  function updateGraphTransform() {
    graph.setAttribute('transform', `translate(${translateX},${translateY}) scale(${scale})`);
  }
  updateGraphTransform();
  svg.addEventListener('wheel', function(evt) {
    evt.preventDefault();
    const delta = evt.deltaY > 0 ? 0.9 : 1.1;
    const newScale = scale * delta;
    if (newScale < 0.2 || newScale > 5) return;
    const rect = svg.getBoundingClientRect();
    const offsetX = evt.clientX - rect.left - translateX;
    const offsetY = evt.clientY - rect.top - translateY;
    translateX -= offsetX * (delta - 1);
    translateY -= offsetY * (delta - 1);
    scale = newScale;
    updateGraphTransform();
  });
  svg.addEventListener('mousedown', function(evt) {
    if (evt.target.closest('g.node')) return;
    panning = true;
    panStartX = evt.clientX;
    panStartY = evt.clientY;
  });
  svg.addEventListener('mousemove', function(evt) {
    if (!panning) return;
    const dx = evt.clientX - panStartX;
    const dy = evt.clientY - panStartY;
    translateX += dx;
    translateY += dy;
    panStartX = evt.clientX;
    panStartY = evt.clientY;
    updateGraphTransform();
  });
  document.addEventListener('mouseup', function(evt) {
    panning = false;
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
/* Fremhev noder ved hover */
.node:hover rect, .node:hover ellipse {{ fill: #f0faff; }}
.node:hover text {{ font-weight: bold; }}
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
<g id="legend" transform="translate(20,70)">
  <rect x="0" y="0" width="20" height="12" rx="2" ry="2" style="fill:#e9ecef;stroke:#495057;stroke-width:1.2"></rect>
  <text x="25" y="10" font-size="12" font-family="Arial, Helvetica, sans-serif" fill="#212529">Selskap</text>
  <ellipse cx="10" cy="25" rx="10" ry="6" style="fill:#d0e6fa;stroke:#5b84ca;stroke-width:1.2"></ellipse>
  <text x="25" y="29" font-size="12" font-family="Arial, Helvetica, sans-serif" fill="#212529">Privatperson</text>
  <line x1="0" y1="42" x2="20" y2="42" style="stroke:#28a745;stroke-width:2;"></line>
  <text x="25" y="45" font-size="12" font-family="Arial, Helvetica, sans-serif" fill="#212529">≥ 50 % eierandel</text>
  <line x1="0" y1="58" x2="20" y2="58" style="stroke:#ffc107;stroke-width:2;"></line>
  <text x="25" y="61" font-size="12" font-family="Arial, Helvetica, sans-serif" fill="#212529">10–49 % eierandel</text>
  <line x1="0" y1="74" x2="20" y2="74" style="stroke:#dc3545;stroke-width:2;"></line>
  <text x="25" y="77" font-size="12" font-family="Arial, Helvetica, sans-serif" fill="#212529">&lt; 10 % eierandel</text>
</g>
<g id="graph">
{''.join(edge_svgs)}
{''.join(node_svgs)}
</g>
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

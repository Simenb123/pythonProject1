"""
org_view.py - Interactive Tkinter canvas for displaying shareholder ownership graphs.

This module defines the ``OrgChartCanvas`` class, a subclass of ``tk.Canvas``
that can draw a graph of companies and people based on data provided by
``OrgChartModel``.  It supports simple hierarchical layout, interactive
drag-and-drop repositioning of nodes, click callbacks for displaying
details, and color-coded edges based on ownership percentage.

The canvas does not itself fetch data; it expects to be given a model
instance populated with nodes and edges.  A controller should
instantiate the canvas, build the model, call ``draw_graph`` to render,
and handle callbacks.
"""
from __future__ import annotations

import tkinter as tk
from tkinter import font as tkfont
from typing import Dict, Tuple, Optional, Callable, Any, List

from .org_model import OrgChartModel, Node, Edge


class OrgChartCanvas(tk.Canvas):
    """
    A canvas widget capable of drawing an ownership graph and supporting
    basic interactivity (click, drag, tooltips, etc.).

    To use this widget:
      1. Create an ``OrgChartModel`` and call ``model.build_graph()``.
      2. Instantiate ``OrgChartCanvas`` with a controller callback.
      3. Call ``draw_graph(model)`` to render the nodes and edges.
      4. Handle node clicks via the ``on_node_click`` callback.

    The canvas uses simple level-based layout: owners are placed above
    their companies and subsidiaries below.  Horizontal spacing is
    automatic based on the number of nodes per level.
    """

    # Default visual constants
    NODE_WIDTH = 140
    NODE_HEIGHT = 50
    X_SPACING = 200
    Y_SPACING = 150
    NODE_COLORS = {
        True: {  # is_company
            "fill": "#EEEEEE",  # light grey
            "outline": "#666666",  # dark grey
        },
        False: {
            "fill": "#D6EAF8",  # light blue
            "outline": "#4682B4",  # steel blue
        },
    }
    ROOT_OUTLINE_WIDTH = 3

    def __init__(self,
                 master: tk.Misc,
                 on_node_click: Optional[Callable[[Node], None]] = None,
                 on_node_double_click: Optional[Callable[[Node], None]] = None,
                 *,
                 read_only: bool = False,
                 **kwargs: Any) -> None:
        super().__init__(master, background="white", highlightthickness=0, **kwargs)
        # Callback invoked when a node is clicked; receives Node instance
        self._on_node_click_callback = on_node_click
        # Callback invoked when a node is double-clicked; receives Node instance
        self._on_node_double_click_callback = on_node_double_click
        # Storage for drawn items: mapping item id -> node id
        self._item_to_node: Dict[int, str] = {}
        # Storage for edges: mapping (owner_id, company_id) -> line item id
        self._edge_items: Dict[Tuple[str, str], int] = {}
        # Dragging state
        self._drag_data: Dict[str, Any] = {}
        # Font for node labels
        self._font = tkfont.Font(family="Helvetica", size=10)
        # Font for edge percentage labels (smaller)
        self._edge_font = tkfont.Font(family="Helvetica", size=8)

        # Storage for edge percentage text items.  Maps (owner_id, company_id)
        # to the text item ID.  Used to update positions when nodes move.
        self._edge_label_items: Dict[Tuple[str, str], int] = {}

        # ------------------------------------------------------------------
        # Selection state
        #
        # ``selected_nodes`` keeps track of which node IDs are currently
        # selected.  Selected nodes are highlighted with a thicker,
        # coloured outline.  When dragging, all selected nodes move
        # together.  The selection can be modified via single click
        # (replace), Shift+click (add), Control+click (toggle), or by
        # drawing a selection rectangle over multiple nodes.
        self.selected_nodes: set[str] = set()
        # During selection rectangle drag, ``_selection_rect_id`` holds
        # the canvas item id of the rectangle, and ``_selection_start``
        # records the starting point in canvas coordinates.  When
        # dragging ends, the rectangle is removed and nodes inside are
        # selected.
        self._selection_rect_id: Optional[int] = None
        self._selection_start: Optional[Tuple[float, float]] = None
        # Maps node id to the shape and text item ids.  This allows
        # highlighting/unhighlighting specific nodes without affecting
        # other elements that share the same tag.
        self._node_shape_id: Dict[str, int] = {}
        self._node_text_id: Dict[str, int] = {}

        # Zoom state; 1.0 is default. Adjust via zoom_in/zoom_out methods
        self._zoom = 1.0

        # Read-only flag: if True, disable dragging and selection interactions.
        self.read_only = read_only

        # Bind Control+MouseWheel for zooming (Ctrl+scroll up/down)
        # Use a generic MouseWheel binding and check the Control modifier bit (0x0004).
        # This allows consistent zoom behaviour across platforms.
        self.bind_all("<MouseWheel>", self._on_mousewheel)
        # Bind right mouse button drag for panning (move the view)
        self.bind("<ButtonPress-3>", self._start_pan)
        self.bind("<B3-Motion>", self._on_pan)

        # Bindings for dragging and selection.  If read_only is True,
        # clicks will still trigger node selection/details, but dragging
        # and rectangle selection are disabled in the handlers themselves.
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<B1-Motion>", self._on_motion)
        self.bind("<ButtonRelease-1>", self._on_release)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def draw_graph(self, model: OrgChartModel) -> None:
        """
        Clear the canvas and draw all nodes and edges described in
        ``model``.  Performs a simple hierarchical layout before
        rendering.
        """
        self.delete("all")
        self._item_to_node.clear()
        self._edge_items.clear()

        # Assign positions to nodes based on hierarchical levels
        levels = self._assign_levels(model)
        positions = self._assign_positions(levels, model)

        # Update model node coordinates for later use (e.g. edge refresh)
        for node_id, (x, y) in positions.items():
            node = model.nodes[node_id]
            node.x, node.y = x, y

        # Store root id on the canvas for convenience (used in drawing root)
        self.root_id = model.root_orgnr

        # Draw edges first (so they appear behind nodes)
        for edge in model.edges:
            self._draw_edge(model, edge)

        # Draw nodes
        for node_id, node in model.nodes.items():
            self._draw_node(node_id, node)

        # Store root id on the canvas for convenience (used in drawing root)
        self.root_id = model.root_orgnr

        # Do not draw legend on the canvas; legend is rendered via controller UI

    # ------------------------------------------------------------------
    # Layout helpers
    # ------------------------------------------------------------------
    def _assign_levels(self, model: OrgChartModel) -> Dict[int, List[str]]:
        """
        Compute a mapping from hierarchical level to list of node IDs.

        The root node is at level 0.  Owners are assigned negative levels,
        and subsidiaries positive levels.  Breadth-first traversal is used
        to ensure a reasonable grouping.  Levels are contiguous integers.
        """
        from collections import deque
        levels: Dict[int, List[str]] = {}
        node_levels: Dict[str, int] = {}

        root_id = model.root_orgnr
        node_levels[root_id] = 0
        queue: deque[str] = deque([root_id])
        while queue:
            current = queue.popleft()
            cur_level = node_levels[current]
            # Owners: go one level up
            for owner, _edge in model.get_owners_of(current):
                if owner.id not in node_levels:
                    node_levels[owner.id] = cur_level - 1
                    queue.append(owner.id)
            # Children: go one level down
            for child, _edge in model.get_children_of(current):
                if child.id not in node_levels:
                    node_levels[child.id] = cur_level + 1
                    queue.append(child.id)
        # Build levels dict
        for node_id, lvl in node_levels.items():
            levels.setdefault(lvl, []).append(node_id)
        return levels

    def _assign_positions(self, levels: Dict[int, List[str]], model: OrgChartModel) -> Dict[str, Tuple[float, float]]:
        """
        Compute a more readable layout for the graph based on hierarchical levels.

        This method uses a simple barycenter heuristic to order nodes within
        each level.  For levels above the root (negative indices), the
        position of each node is determined by the average x-coordinate
        of its children at the level below.  For levels below the root,
        the ordering is determined by the average x-coordinate of its
        owners at the level above.  If a node has no neighbours in the
        adjacent level, it is positioned at the centre of the level.

        Each level is centred horizontally around x=0, with nodes
        separated by ``X_SPACING``.  The y-coordinate is determined by
        the level index relative to the minimum level and ``Y_SPACING``.
        """
        positions: Dict[str, Tuple[float, float]] = {}
        if not levels:
            return positions

        min_level = min(levels.keys())
        max_level = max(levels.keys())

        # First place the root level (0).  If multiple roots exist,
        # spread them evenly.
        root_level = levels.get(0, [])
        if root_level:
            n = len(root_level)
            width = (n - 1) * self.X_SPACING if n > 1 else 0
            for idx, node_id in enumerate(sorted(root_level)):
                x = idx * self.X_SPACING - width / 2
                y = (0 - min_level) * self.Y_SPACING
                positions[node_id] = (x, y)

        # Assign positions for levels below the root (positive levels)
        for lvl in range(1, max_level + 1):
            node_ids = levels.get(lvl, [])
            if not node_ids:
                continue
            # Compute barycentres from owners (level-1) positions
            barycentres: List[Tuple[str, float]] = []
            for nid in node_ids:
                # owners at level-1
                neighbours = [owner.id for owner, _ in model.get_owners_of(nid)]
                xs = [positions[owner_id][0] for owner_id in neighbours if owner_id in positions]
                if xs:
                    b = sum(xs) / len(xs)
                else:
                    b = 0.0
                barycentres.append((nid, b))
            # Sort nodes by barycentre
            barycentres.sort(key=lambda t: t[1])
            n = len(barycentres)
            width = (n - 1) * self.X_SPACING if n > 1 else 0
            for idx, (nid, _b) in enumerate(barycentres):
                x = idx * self.X_SPACING - width / 2
                y = (lvl - min_level) * self.Y_SPACING
                positions[nid] = (x, y)

        # Assign positions for levels above the root (negative levels)
        # We iterate from -1 down to min_level
        for lvl in range(-1, min_level - 1, -1):
            node_ids = levels.get(lvl, [])
            if not node_ids:
                continue
            # Compute barycentres from children (level+1) positions
            barycentres: List[Tuple[str, float]] = []
            for nid in node_ids:
                # children at level+1
                neighbours = [child.id for child, _ in model.get_children_of(nid)]
                xs = [positions[child_id][0] for child_id in neighbours if child_id in positions]
                if xs:
                    b = sum(xs) / len(xs)
                else:
                    b = 0.0
                barycentres.append((nid, b))
            # Sort nodes by barycentre
            barycentres.sort(key=lambda t: t[1])
            n = len(barycentres)
            width = (n - 1) * self.X_SPACING if n > 1 else 0
            for idx, (nid, _b) in enumerate(barycentres):
                x = idx * self.X_SPACING - width / 2
                y = (lvl - min_level) * self.Y_SPACING
                positions[nid] = (x, y)

        return positions

    # ------------------------------------------------------------------
    # Drawing helpers
    # ------------------------------------------------------------------
    def _draw_node(self, node_id: str, node: Node) -> None:
        """Draw a single node (rectangle or oval) with text labels.

        Node boxes are sized dynamically based on the text content.  The name
        and id are split onto two lines, measured with the canvas font, and
        padding is added.  For the root node, we apply a thicker outline
        and a distinct fill colour to highlight it.  The computed width
        and height are stored on the Node instance (as ``canvas_w`` and
        ``canvas_h``) for use when updating edges during drag.
        """
        x, y = node.x, node.y
        # Build label text: name and id on separate lines
        name_part = node.name or node.id
        id_disp = f"({node.id})" if node.id else ""
        if id_disp:
            label_lines = f"{name_part}\n{id_disp}"
        else:
            label_lines = name_part
        # Measure width of each line using the font
        lines = label_lines.split("\n")
        widths = [self._font.measure(line) for line in lines]
        # Use a minimum width for tiny names
        text_width = max(widths) if widths else 0
        # Horizontal padding
        pad_x = 16
        node_w = max(self.NODE_WIDTH, text_width + pad_x)
        # Height: number of lines * line height + vertical padding
        line_height = self._font.metrics("linespace")
        node_h = max(self.NODE_HEIGHT, line_height * len(lines) + 10)
        # Store computed dimensions on node for later use (edge updates)
        node.canvas_w = node_w  # type: ignore[attr-defined]
        node.canvas_h = node_h  # type: ignore[attr-defined]
        # Determine box coordinates
        x0 = x - node_w / 2
        y0 = y - node_h / 2
        x1 = x + node_w / 2
        y1 = y + node_h / 2
        # Determine colours
        color_cfg = self.NODE_COLORS[node.is_company]
        # Root detection: compare to stored root_id if available
        is_root = hasattr(self, "root_id") and node.id == getattr(self, "root_id")
        # Adjust outline and fill for root highlight
        if is_root:
            outline_width = self.ROOT_OUTLINE_WIDTH
            # Use a pale yellow fill to distinguish root
            fill_color = "#FFF9C4"
            outline_color = "#666666"
        else:
            outline_width = 1
            fill_color = color_cfg["fill"]
            outline_color = color_cfg["outline"]
        # Draw shape (rectangle for company, oval for person)
        if node.is_company:
            shape_id = self.create_rectangle(
                x0, y0, x1, y1,
                fill=fill_color,
                outline=outline_color,
                width=outline_width,
                tags=("node", f"node_{node_id}")
            )
        else:
            shape_id = self.create_oval(
                x0, y0, x1, y1,
                fill=fill_color,
                outline=outline_color,
                width=outline_width,
                tags=("node", f"node_{node_id}")
            )
        # Draw label centered in the box
        text_id = self.create_text(
            x, y,
            text=label_lines,
            font=self._font,
            fill="#333333",
            justify="center",
            tags=("node", f"node_{node_id}")
        )
        # Map both shape and text items to node id for lookup
        self._item_to_node[shape_id] = node_id
        self._item_to_node[text_id] = node_id
        # Store shape/text ids for highlighting and selection
        self._node_shape_id[node_id] = shape_id
        self._node_text_id[node_id] = text_id

        # Bind double-click event to trigger a callback if provided
        if self._on_node_double_click_callback:
            # Bind to the tag so both shape and text trigger the event
            tag = f"node_{node_id}"
            self.tag_bind(tag, "<Double-Button-1>", lambda event, nid=node_id: self._handle_double_click(event, nid))

    def _handle_double_click(self, event: tk.Event, node_id: str) -> None:
        """Internal handler for node double-click.  Invokes the controller callback."""
        if not self._on_node_double_click_callback:
            return
        # Look up Node object from model if available
        node_obj: Optional[Node] = None
        if hasattr(self, "model"):
            node_obj = self.model.nodes.get(node_id)  # type: ignore[attr-defined]
        # Fall back to an empty object if not found
        if node_obj is None:
            return
        # Invoke the callback with the Node
        self._on_node_double_click_callback(node_obj)

    def _draw_edge(self, model: OrgChartModel, edge: Edge) -> None:
        """Draw a line with an arrow between owner and company nodes."""
        owner = model.nodes.get(edge.owner_id)
        company = model.nodes.get(edge.company_id)
        if not owner or not company:
            return
        # Determine start and end positions (center to center).
        # Coordinates in the model (node.x, node.y) are defined in a
        # logical coordinate system (unscaled).  When the canvas is
        # zoomed, all items are scaled, but the model coordinates are
        # not updated.  To align edges correctly after zooming, multiply
        # the model coordinates by the current zoom factor when
        # computing the anchor points.
        x1, y1 = owner.x * self._zoom, owner.y * self._zoom
        x2, y2 = company.x * self._zoom, company.y * self._zoom
        # Adjust start/end to edge of the shapes (approximate)
        # If owner is above company, we start at bottom of owner, end at top of company
        dx = x2 - x1
        dy = y2 - y1
        # Determine orientation for bounding boxes
        def adjust(x: float, y: float, node: Node, direction: int) -> Tuple[float, float]:
            """Return the point on the edge of a node in the given direction.

            Multiply the half-height by the zoom factor so offsets scale
            with the canvas.  ``direction`` is +1 for the bottom edge
            (when the line goes downwards) and -1 for the top edge.
            """
            half_h = getattr(node, "canvas_h", self.NODE_HEIGHT) / 2.0
            return (x, y + half_h * direction * self._zoom)

        if dy >= 0:
            # owner above or on same level: start bottom of owner, end top of company
            start = adjust(x1, y1, owner, +1)
            end = adjust(x2, y2, company, -1)
        else:
            # owner below (rare but handle) – reverse
            start = adjust(x1, y1, owner, -1)
            end = adjust(x2, y2, company, +1)

        # Determine line colour and width based on ownership percentage
        pct = edge.ownership_pct
        if pct is None:
            colour = "#888888"
            width = 1
        else:
            if pct >= 50:
                colour = "#3CB371"  # medium green
            elif pct >= 10:
                colour = "#F4D03F"  # golden yellow
            else:
                colour = "#E74C3C"  # red
            width = 1 + (pct / 50 if pct < 50 else 2)

        line_id = self.create_line(
            start[0], start[1], end[0], end[1],
            arrow=tk.LAST,
            arrowshape=(8, 10, 4),
            fill=colour,
            width=width,
            smooth=False,
        )
        self._edge_items[(edge.owner_id, edge.company_id)] = line_id

        # Draw percentage label at the midpoint of the edge if available
        if pct is not None:
            try:
                pct_val = float(pct)
            except Exception:
                pct_val = None
            if pct_val is not None:
                label = f"{pct_val:.2f}%"
                # Compute midpoint and offset slightly above the line.
                # Apply zoom factor to the vertical offset so labels stay
                # close to the line when zooming.
                mid_x = (start[0] + end[0]) / 2.0
                mid_y = (start[1] + end[1]) / 2.0 - 8.0 * self._zoom
                label_id = self.create_text(
                    mid_x, mid_y,
                    text=label,
                    font=self._edge_font,
                    fill=colour,
                    tags=("edge_label",)
                )
                # Store the label id so we can update its position later
                self._edge_label_items[(edge.owner_id, edge.company_id)] = label_id

    def _draw_legend(self) -> None:
        """Draw a simple legend explaining colours and shapes."""
        # Use small font for legend
        legend_font = tkfont.Font(family="Helvetica", size=8)
        x, y = 10, 10
        # Company box
        self.create_rectangle(x, y, x + 16, y + 12, fill=self.NODE_COLORS[True]["fill"],
                              outline=self.NODE_COLORS[True]["outline"], width=1)
        self.create_text(x + 20, y + 6, text="Selskap", anchor="w", font=legend_font)
        y += 18
        # Person oval
        self.create_oval(x, y, x + 16, y + 12, fill=self.NODE_COLORS[False]["fill"],
                         outline=self.NODE_COLORS[False]["outline"], width=1)
        self.create_text(x + 20, y + 6, text="Privatperson", anchor="w", font=legend_font)
        y += 20
        # Colour bars for edge percentages
        self.create_line(x, y, x + 16, y, fill="#3CB371", width=3)
        self.create_text(x + 20, y, text="≥ 50 % eierandel", anchor="w", font=legend_font)
        y += 14
        self.create_line(x, y, x + 16, y, fill="#F4D03F", width=3)
        self.create_text(x + 20, y, text="10–49 % eierandel", anchor="w", font=legend_font)
        y += 14
        self.create_line(x, y, x + 16, y, fill="#E74C3C", width=3)
        self.create_text(x + 20, y, text="< 10 % eierandel", anchor="w", font=legend_font)

    # ------------------------------------------------------------------
    # Zoom handling
    # ------------------------------------------------------------------
    def _zoom_all(self, factor: float) -> None:
        """
        Scale all canvas items by ``factor`` around the origin.  Updates the
        internal zoom level and adjusts scrollregion.  Limits zoom to
        reasonable values (between 0.2x and 5x).
        """
        new_zoom = self._zoom * factor
        # Clamp zoom factor to [0.2, 5]
        if new_zoom < 0.2 or new_zoom > 5.0:
            return
        self._zoom = new_zoom
        # Scale all items around the origin (0,0)
        self.scale("all", 0, 0, factor, factor)
        # Adjust scrollregion to new bounding box
        bbox = self.bbox("all")
        if bbox:
            self.configure(scrollregion=bbox)

    def zoom_in(self) -> None:
        """Zoom in by 10%."""
        self._zoom_all(1.1)

    def zoom_out(self) -> None:
        """Zoom out by 10%."""
        self._zoom_all(0.9)

    # ------------------------------------------------------------------
    # MouseWheel and panning handlers
    # ------------------------------------------------------------------
    def _on_mousewheel(self, event: tk.Event) -> None:
        """
        Respond to mouse wheel events.  If the Control key is held, zoom
        in or out.  Otherwise, default scrolling behaviour applies.
        """
        # Modifier state: 0x0004 is Control on most platforms
        if (event.state & 0x0004) != 0:
            # Determine zoom factor based on scroll direction
            factor = 1.1 if event.delta > 0 else 0.9
            self._zoom_all(factor)
            # Prevent default scrolling when zooming
            return "break"
        # Otherwise, allow default behaviour for scrolling
        return None

    def _start_pan(self, event: tk.Event) -> None:
        """Start panning the canvas (right button pressed)."""
        self.scan_mark(event.x, event.y)

    def _on_pan(self, event: tk.Event) -> None:
        """Handle mouse movement while right button is held for panning."""
        # gain=1 makes panning follow the mouse movement
        self.scan_dragto(event.x, event.y, gain=1)

    # ------------------------------------------------------------------
    # Event handling
    # ------------------------------------------------------------------
    def _on_press(self, event: tk.Event) -> None:
        """Handle mouse button press for potential dragging or clicking."""
        # Cancel any ongoing selection rectangle
        if self._selection_rect_id is not None:
            self.delete(self._selection_rect_id)
            self._selection_rect_id = None
            self._selection_start = None
        # Determine if the click is on a node item
        item = self.find_withtag("current")
        if item:
            item_id = item[0]
            node_id = self._item_to_node.get(item_id)
            if node_id:
                # In read-only mode: treat this as a simple click.  Do not allow
                # selection rectangle or dragging.  Just highlight and invoke callback.
                if self.read_only:
                    # Clear selection and highlight this node
                    self._clear_selection()
                    self.selected_nodes.add(node_id)
                    self._highlight_node(node_id)
                    # Invoke click callback if provided
                    if self._on_node_click_callback and hasattr(self, "model"):
                        node_obj = self.model.nodes.get(node_id)  # type: ignore[attr-defined]
                        if node_obj:
                            self._on_node_click_callback(node_obj)
                    return
                # Not read-only: handle selection and start drag as usual
                # Modifier keys: Shift (0x0001), Control (0x0004)
                shift_pressed = (event.state & 0x0001) != 0
                ctrl_pressed = (event.state & 0x0004) != 0
                # Update selection based on modifiers
                if node_id not in self.selected_nodes:
                    if not shift_pressed and not ctrl_pressed:
                        # Clear existing selection and select this node
                        self._clear_selection()
                    # Add to selection
                    self.selected_nodes.add(node_id)
                    self._highlight_node(node_id)
                else:
                    # Node already selected
                    if ctrl_pressed:
                        # Toggle off selection
                        self._unhighlight_node(node_id)
                        self.selected_nodes.discard(node_id)
                    # If shift pressed, do nothing (keep it selected)
                # Start dragging; record initial position
                self._drag_data = {
                    "node_id": node_id,
                    "x": event.x,
                    "y": event.y,
                }
                return
        # If not clicking on a node, start selection rectangle
        # Clear previous selection unless Shift is held (add to selection)
        shift_pressed = (event.state & 0x0001) != 0
        if not shift_pressed:
            self._clear_selection()
        # In read-only mode, do not start selection rectangle
        if self.read_only:
            return
        # Record starting point in canvas coordinates
        self._selection_start = (self.canvasx(event.x), self.canvasy(event.y))
        # Draw rectangle (initial zero size).  Use dashed outline for clarity.
        self._selection_rect_id = self.create_rectangle(
            self._selection_start[0], self._selection_start[1],
            self._selection_start[0], self._selection_start[1],
            outline="#4A90E2",
            dash=(4, 2),
            width=1
        )

    def _on_motion(self, event: tk.Event) -> None:
        """Handle mouse movement while button 1 is held (dragging)."""
        # Ignore dragging if in read-only mode
        if self.read_only:
            return
        # If selection rectangle is active, update its size
        if self._selection_rect_id is not None and self._selection_start is not None:
            # Update rectangle coordinates based on drag
            x0, y0 = self._selection_start
            x1, y1 = self.canvasx(event.x), self.canvasy(event.y)
            self.coords(self._selection_rect_id, x0, y0, x1, y1)
            return
        # If dragging nodes, move all selected nodes together
        if not self._drag_data:
            return
        # Compute deltas
        dx = event.x - self._drag_data.get("x", event.x)
        dy = event.y - self._drag_data.get("y", event.y)
        self._drag_data["x"] = event.x
        self._drag_data["y"] = event.y
        if not self.selected_nodes:
            return
        # Move each selected node and update model coordinates
        if hasattr(self, "model"):
            for nid in list(self.selected_nodes):
                tag = f"node_{nid}"
                self.move(tag, dx, dy)
                node = self.model.nodes.get(nid)  # type: ignore[attr-defined]
                if node:
                    # adjust by zoom
                    if self._zoom != 0:
                        node.x += dx / self._zoom
                        node.y += dy / self._zoom
                    else:
                        node.x += dx
                        node.y += dy
            # Update edges connected to any moved nodes
            for edge in self.model.edges:
                if edge.owner_id in self.selected_nodes or edge.company_id in self.selected_nodes:
                    line_id = self._edge_items.get((edge.owner_id, edge.company_id))
                    if line_id:
                        self._update_edge_coords(edge, line_id)

    def _on_release(self, event: tk.Event) -> None:
        """Handle mouse button release (end dragging or click)."""
        # In read-only mode, ignore release events (no dragging or selection)
        if self.read_only:
            # Clear drag state and selection rectangle if any
            if self._selection_rect_id is not None:
                self.delete(self._selection_rect_id)
                self._selection_rect_id = None
                self._selection_start = None
            self._drag_data = {}
            return
        # If a selection rectangle was being dragged, finalize the selection
        if self._selection_rect_id is not None and self._selection_start is not None:
            # Get rectangle coords (x0,y0,x1,y1) in canvas coordinates
            coords = self.coords(self._selection_rect_id)
            if coords:
                x0, y0, x1, y1 = coords
                # Ensure ordering
                x_min, x_max = (x0, x1) if x0 < x1 else (x1, x0)
                y_min, y_max = (y0, y1) if y0 < y1 else (y1, y0)
                # Clear current selection before adding new ones
                self._clear_selection()
                # Iterate over all nodes and select those fully within the rectangle
                if hasattr(self, "model"):
                    for nid, node in self.model.nodes.items():  # type: ignore[attr-defined]
                        shape_id = self._node_shape_id.get(nid)
                        if shape_id:
                            bbox = self.bbox(shape_id)
                            if bbox:
                                sx0, sy0, sx1, sy1 = bbox
                                # Check if node bounding box is within selection rectangle
                                if sx0 >= x_min and sy0 >= y_min and sx1 <= x_max and sy1 <= y_max:
                                    self.selected_nodes.add(nid)
                                    self._highlight_node(nid)
            # Remove selection rectangle
            self.delete(self._selection_rect_id)
            self._selection_rect_id = None
            self._selection_start = None
            # Reset drag data and return; no further click handling
            self._drag_data = {}
            return
        # Otherwise handle end of drag or simple click
        if not self._drag_data:
            return
        node_id = self._drag_data.get("node_id")
        # Clear drag state
        self._drag_data = {}
        if not node_id:
            return
        # Trigger click callback if appropriate
        item = self.find_withtag("current")
        if item:
            item_id = item[0]
            clicked_node_id = self._item_to_node.get(item_id)
            if clicked_node_id and self._on_node_click_callback:
                node_obj = self.model.nodes.get(clicked_node_id) if hasattr(self, "model") else None
                if node_obj:
                    self._on_node_click_callback(node_obj)

    def _update_edge_coords(self, edge: Edge, line_id: int) -> None:
        """
        Update the coordinates of an existing edge line based on the current
        node positions.  This is called during drag operations.
        """
        owner = self.model.nodes.get(edge.owner_id)  # type: ignore[attr-defined]
        company = self.model.nodes.get(edge.company_id)  # type: ignore[attr-defined]
        if not owner or not company:
            return
        # Recompute adjusted start/end points
        def adjust(x: float, y: float, node: Node, direction: int) -> Tuple[float, float]:
            """Return a point on the edge of ``node`` based on direction.

            Multiply the half-height by the zoom factor so offsets scale
            with the canvas.  Uses scaled coordinates (node.x*zoom, node.y*zoom)
            for correct positioning after zoom.
            """
            half_h = getattr(node, "canvas_h", self.NODE_HEIGHT) / 2.0
            return (x, y + half_h * direction * self._zoom)
        # Use scaled coordinates for start and end positions
        owner_scaled_x = owner.x * self._zoom
        owner_scaled_y = owner.y * self._zoom
        company_scaled_x = company.x * self._zoom
        company_scaled_y = company.y * self._zoom
        if company_scaled_y >= owner_scaled_y:
            start = adjust(owner_scaled_x, owner_scaled_y, owner, +1)
            end = adjust(company_scaled_x, company_scaled_y, company, -1)
        else:
            start = adjust(owner_scaled_x, owner_scaled_y, owner, -1)
            end = adjust(company_scaled_x, company_scaled_y, company, +1)
        self.coords(line_id, start[0], start[1], end[0], end[1])

        # Update label position if it exists
        label_id = self._edge_label_items.get((edge.owner_id, edge.company_id))
        if label_id:
            mid_x = (start[0] + end[0]) / 2.0
            # Apply zoom factor to vertical offset to keep label near the line
            mid_y = (start[1] + end[1]) / 2.0 - 8.0 * self._zoom
            self.coords(label_id, mid_x, mid_y)

    # Helper to associate canvas with its model (set by controller)
    def set_model(self, model: OrgChartModel) -> None:
        """Store a reference to the model for use during drag operations."""
        self.model = model  # type: ignore[attr-defined]

    # ------------------------------------------------------------------
    # Selection and highlighting helpers
    # ------------------------------------------------------------------
    def _clear_selection(self) -> None:
        """Remove selection highlight from all selected nodes and clear the set."""
        for nid in list(self.selected_nodes):
            self._unhighlight_node(nid)
        self.selected_nodes.clear()

    def _highlight_node(self, node_id: str) -> None:
        """Visually highlight a node by altering its outline colour and width."""
        shape_id = self._node_shape_id.get(node_id)
        if shape_id:
            # Use royal blue for highlight; maintain current fill
            self.itemconfig(shape_id, outline="#4169E1", width=2)

    def _unhighlight_node(self, node_id: str) -> None:
        """Remove highlight from a node, restoring its default colours."""
        shape_id = self._node_shape_id.get(node_id)
        if not shape_id:
            return
        # Determine default colours based on node type and root status
        node = None
        if hasattr(self, "model"):
            node = self.model.nodes.get(node_id)  # type: ignore[attr-defined]
        if node is not None:
            color_cfg = self.NODE_COLORS[node.is_company]
            is_root = hasattr(self, "root_id") and node.id == getattr(self, "root_id")
            outline_width = self.ROOT_OUTLINE_WIDTH if is_root else 1
            outline_color = color_cfg["outline"]
            fill_color = "#FFF9C4" if is_root else color_cfg["fill"]
            self.itemconfig(shape_id, outline=outline_color, width=outline_width, fill=fill_color)

"""
org_controller.py - Tkinter application/controller for interactive ownership graphs.

This module ties together the `OrgChartModel` (data layer) and
`OrgChartCanvas` (presentation layer) to provide a simple user
interface for searching companies in the aksjonærregister, building
graphs, and interacting with them.  It allows a user to search by
company name or organisation number, select a company, and visualise
its ownership structure.  Clicking on any node in the graph will
display details about that company or person.  The application does
not modify the underlying aksjonærregister; any edits made by the
user (e.g. adding new ownership relations) are stored in memory only.

This controller is designed to be self-contained and can be run from
the command line or imported into another Tkinter-based program.  It
depends only on standard Python libraries and the modules in this
package (db.py, org_model.py and org_view.py).
"""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Tuple

# Import modules from the same directory.  We avoid relative imports here
# so that this module can be run as a script without requiring a package
# context.  The modules ``db``, ``org_model`` and ``org_view`` live in
# the project root alongside this file.
# Import modules relative to this file if part of a package; fall back to
# absolute imports when running as a script.  This allows the controller
# to be executed both via ``python -m app.aksjonærregister.run_orgchart``
# and via ``python org_controller.py``.
try:
    from . import db  # type: ignore[import-not-found]
    from .org_model import OrgChartModel, Node  # type: ignore[import-not-found]
    from .org_view import OrgChartCanvas  # type: ignore[import-not-found]
except ImportError:
    import db  # type: ignore[import-not-found]
    from org_model import OrgChartModel, Node  # type: ignore[import-not-found]
    from org_view import OrgChartCanvas  # type: ignore[import-not-found]


class OrgChartApp(tk.Tk):
    """
    A simple Tkinter application for exploring ownership structures.

    The app presents a search bar for looking up companies by name or
    orgnr, a list of matching companies, and an interactive canvas
    showing the ownership graph for the selected company.  A details
    pane shows information about the currently selected node in the
    graph.  Users can drag nodes to reposition them for clarity.

    The application can be extended with editing features (adding
    ownership relations) by updating the model and redrawing the
    canvas.  Saving and loading of user modifications can be added
    later via JSON (see OrgChartModel for suggestions).
    """

    def __init__(self) -> None:
        super().__init__()
        self.title("Aksjonærregister – interaktiv orgkart (Tkinter)")
        # Default window size; user can resize
        self.geometry("1200x800")

        # Open database connection (read-only for search)
        self.conn = db.open_conn()

        # Currently selected company orgnr (root)
        self.current_orgnr: Optional[str] = None
        # Current OrgChartModel
        self.model: Optional[OrgChartModel] = None
        # Build UI components
        self._build_ui()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        """Construct all widgets and layout."""
        # Top frame: search controls
        search_frame = ttk.Frame(self, padding=6)
        search_frame.pack(fill=tk.X, expand=False)
        ttk.Label(search_frame, text="Søk:" ).pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40)
        entry.pack(side=tk.LEFT, padx=4)
        entry.bind("<Return>", lambda _e: self._do_search())
        # Radio buttons to choose search by name/orgnr
        self.search_by = tk.StringVar(value="navn")
        ttk.Radiobutton(search_frame, text="Navn", variable=self.search_by, value="navn").pack(side=tk.LEFT)
        ttk.Radiobutton(search_frame, text="Orgnr", variable=self.search_by, value="orgnr").pack(side=tk.LEFT)
        ttk.Button(search_frame, text="Søk", command=self._do_search).pack(side=tk.LEFT, padx=4)

        # Toggle button to collapse/expand the left panel. When collapsed,
        # the graph takes full width; when expanded, the search/results
        # panel is visible.  Use a simple arrow for the button label that
        # flips depending on state.  See `_toggle_left_panel` for logic.
        self.left_collapsed = False
        self.toggle_btn = ttk.Button(
            search_frame,
            text="◀",  # black left-pointing triangle
            width=2,
            command=self._toggle_left_panel,
        )
        self.toggle_btn.pack(side=tk.RIGHT, padx=(4, 0))

        # Main content: horizontally split into list and canvas+details
        content = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        content.pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

        # Left side: list of companies and owners table
        left_frame = ttk.Frame(content)
        content.add(left_frame, weight=1)

        # Treeview for search results
        self.result_tree = ttk.Treeview(left_frame, columns=("orgnr", "name"), show="headings", height=10)
        self.result_tree.heading("orgnr", text="Orgnr")
        self.result_tree.heading("name", text="Selskap")
        self.result_tree.column("orgnr", width=100, anchor=tk.W)
        self.result_tree.column("name", width=300, anchor=tk.W)
        self.result_tree.pack(fill=tk.BOTH, expand=False, pady=(0, 6))
        self.result_tree.bind("<<TreeviewSelect>>", self._on_select_company)
        # Also bind double-click on result tree items to immediately trigger selection
        self.result_tree.bind("<Double-1>", lambda e: self._on_select_company())

        # Depth controls: allow user to choose how many levels up/down to traverse
        depth_frame = ttk.Frame(left_frame)
        depth_frame.pack(fill=tk.X, pady=(4, 4))
        ttk.Label(depth_frame, text="Dybde opp:").pack(side=tk.LEFT)
        # 0 means unlimited
        self.depth_up_var = tk.IntVar(value=0)
        up_spin = ttk.Spinbox(depth_frame, from_=0, to=8, width=3, textvariable=self.depth_up_var)
        up_spin.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(depth_frame, text="Dybde ned:").pack(side=tk.LEFT)
        self.depth_down_var = tk.IntVar(value=0)
        down_spin = ttk.Spinbox(depth_frame, from_=0, to=8, width=3, textvariable=self.depth_down_var)
        down_spin.pack(side=tk.LEFT)

        # Slider/spinbox for minimum ownership percentage filtering
        thresh_frame = ttk.Frame(left_frame)
        thresh_frame.pack(fill=tk.X, pady=(2, 4))
        ttk.Label(thresh_frame, text="Min eierandel %:").pack(side=tk.LEFT)
        # Using a DoubleVar; values between 0 and 10 (percentage), step 0.5
        self.min_pct_var = tk.DoubleVar(value=0.0)
        self.min_pct_spin = ttk.Spinbox(
            thresh_frame,
            from_=0.0,
            to=100.0,
            increment=0.5,
            width=6,
            textvariable=self.min_pct_var,
            format="%.1f",
        )
        self.min_pct_spin.pack(side=tk.LEFT, padx=(4, 0))
        # When the value changes, rebuild the graph for the current company
        self.min_pct_var.trace_add("write", lambda *_args: self._on_min_pct_change())

        # Owners table for the selected company
        owners_label = ttk.Label(left_frame, text="Eiere i valgt selskap", font=("Helvetica", 10, "bold"))
        owners_label.pack(anchor="w")
        self.owners_tree = ttk.Treeview(
            left_frame,
            columns=("owner_orgnr", "owner_name", "share_class", "owner_country", "owner_zip_place", "shares_owner_num", "shares_company_num", "ownership_pct"),
            show="headings",
            height=12,
        )
        headings = [
            "Eier orgnr/fødselsår",
            "Eier navn",
            "Aksjeklasse",
            "Landkode",
            "Postnr/sted",
            "Antall aksjer",
            "Antall aksjer selskap",
            "Eierandel %",
        ]
        for col, head in zip(self.owners_tree["columns"], headings):
            self.owners_tree.heading(col, text=head)
            # Set column widths; numeric align right
            if col in ("shares_owner_num", "shares_company_num", "ownership_pct"):
                self.owners_tree.column(col, width=100, anchor=tk.E)
            else:
                self.owners_tree.column(col, width=150, anchor=tk.W)
        # Scrollbars for owners table
        owners_scroll_y = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.owners_tree.yview)
        owners_scroll_x = ttk.Scrollbar(left_frame, orient=tk.HORIZONTAL, command=self.owners_tree.xview)
        self.owners_tree.configure(yscrollcommand=owners_scroll_y.set, xscrollcommand=owners_scroll_x.set)
        self.owners_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        owners_scroll_y.pack(fill=tk.Y, side=tk.RIGHT)
        owners_scroll_x.pack(fill=tk.X, side=tk.BOTTOM)

        # Right side: canvas and details panel
        right_frame = ttk.Frame(content)
        content.add(right_frame, weight=2)

        # Canvas for drawing the graph
        self.canvas = OrgChartCanvas(
            right_frame,
            on_node_click=self._on_node_click,
            on_node_double_click=self._on_node_double_click,
        )
        # Put canvas in a scrollable frame to allow panning
        canvas_scroll_y = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        canvas_scroll_x = ttk.Scrollbar(right_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=canvas_scroll_y.set, xscrollcommand=canvas_scroll_x.set)
        # Configure scroll region later when drawing
        self.canvas.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        canvas_scroll_y.pack(fill=tk.Y, side=tk.RIGHT)
        canvas_scroll_x.pack(fill=tk.X, side=tk.BOTTOM)
        # Zoom controls (above legend and details)
        zoom_frame = ttk.Frame(right_frame)
        zoom_frame.pack(fill=tk.X, pady=(4, 0))
        # Use Unicode plus/minus symbols for compact zoom buttons
        # Plus button to zoom in
        ttk.Button(zoom_frame, text="＋", width=3, command=self.canvas.zoom_in).pack(side=tk.LEFT, padx=2)
        # Minus button to zoom out
        ttk.Button(zoom_frame, text="－", width=3, command=self.canvas.zoom_out).pack(side=tk.LEFT)
        # Refresh button to re-draw the graph and reset zoom.  Use a circular arrow symbol.
        ttk.Button(zoom_frame, text="⟳", width=3, command=self._refresh_graph).pack(side=tk.LEFT, padx=(10, 0))

        # Legend explaining shapes and line colours
        legend_frame = ttk.Frame(right_frame)
        legend_frame.pack(fill=tk.X, pady=(4, 0), anchor="w")
        # Create a small canvas to draw legend shapes
        # Allocate a bit more width to avoid line wrapping of legend text
        legend_canvas = tk.Canvas(legend_frame, width=220, height=70, highlightthickness=0, bg="white")
        legend_canvas.pack(side=tk.LEFT)
        # Draw company box
        x, y = 5, 5
        comp_fill = self.canvas.NODE_COLORS[True]["fill"]
        comp_out = self.canvas.NODE_COLORS[True]["outline"]
        person_fill = self.canvas.NODE_COLORS[False]["fill"]
        person_out = self.canvas.NODE_COLORS[False]["outline"]
        legend_canvas.create_rectangle(x, y, x + 16, y + 12, fill=comp_fill, outline=comp_out)
        legend_canvas.create_text(x + 20, y + 6, text="Selskap", anchor="w", font=("Helvetica", 8))
        y += 18
        legend_canvas.create_oval(x, y, x + 16, y + 12, fill=person_fill, outline=person_out)
        legend_canvas.create_text(x + 20, y + 6, text="Privatperson", anchor="w", font=("Helvetica", 8))
        y += 20
        # Colour bars for edge percentages
        legend_canvas.create_line(5, y, 21, y, fill="#3CB371", width=3)
        legend_canvas.create_text(24, y, text="≥ 50 % eierandel", anchor="w", font=("Helvetica", 8))
        y += 14
        legend_canvas.create_line(5, y, 21, y, fill="#F4D03F", width=3)
        legend_canvas.create_text(24, y, text="10–49 % eierandel", anchor="w", font=("Helvetica", 8))
        y += 14
        legend_canvas.create_line(5, y, 21, y, fill="#E74C3C", width=3)
        legend_canvas.create_text(24, y, text="< 10 % eierandel", anchor="w", font=("Helvetica", 8))

        # Details panel below legend
        details_frame = ttk.Labelframe(right_frame, text="Detaljer", padding=6)
        details_frame.pack(fill=tk.X, expand=False, pady=(4, 0))
        # Labels for details (we'll update text when node clicked)
        self.detail_labels = {}
        for field in ["Navn", "ID", "Aksjeklasse", "Landkode", "Postnr/sted", "Antall aksjer", "Antall aksjer selskap", "Eierandel %"]:
            lbl = ttk.Label(details_frame, text=f"{field}: ")
            lbl.pack(anchor="w")
            self.detail_labels[field] = lbl

        # Store frames for toggling visibility later
        self.left_frame = left_frame
        self.right_frame = right_frame
        self.content = content


    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------
    def _do_search(self) -> None:
        """Search for companies and populate the result tree."""
        term = self.search_var.get().strip()
        if not term:
            # Clear results
            for item in self.result_tree.get_children():
                self.result_tree.delete(item)
            return
        try:
            results = db.search_companies(self.conn, term, self.search_by.get())
        except Exception as exc:
            messagebox.showerror("Søkefeil", str(exc))
            return
        # Populate result tree
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        for orgnr, name in results:
            self.result_tree.insert("", tk.END, values=(orgnr, name))

    def _on_select_company(self, _event=None) -> None:
        """Load selected company owners and draw its graph."""
        selection = self.result_tree.selection()
        if not selection:
            return
        orgnr, name = self.result_tree.item(selection[0], "values")
        self.current_orgnr = orgnr
        # Load owners into owners table (from DB via get_owners_full)
        self._load_owners(orgnr)
        # Build model and draw graph
        self._build_and_draw(orgnr)

    def _load_owners(self, company_orgnr: str) -> None:
        """Fetch detailed owner rows from DB and populate owners_tree."""
        try:
            rows = db.get_owners_full(self.conn, company_orgnr)
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke hente eiere: {exc}")
            return
        for item in self.owners_tree.get_children():
            self.owners_tree.delete(item)
        for row in rows:
            # Format numeric fields for display
            owner_orgnr, owner_name, share_class, owner_country, owner_zip_place, shares_owner_num, shares_company_num, pct = row
            def fmt_shares(val):
                try:
                    if val is None:
                        return ""
                    f = float(val)
                    if abs(f - round(f)) < 0.005:
                        return f"{int(round(f)):,}".replace(",", " ")
                    return f"{f:,.2f}".replace(",", " ")
                except Exception:
                    return str(val) if val is not None else ""
            def fmt_pct(val):
                try:
                    return "" if val is None else f"{float(val):.2f}"
                except Exception:
                    return str(val) if val is not None else ""
            values = (
                owner_orgnr or "",
                owner_name or "",
                share_class or "",
                owner_country or "",
                owner_zip_place or "",
                fmt_shares(shares_owner_num),
                fmt_shares(shares_company_num),
                fmt_pct(pct),
            )
            self.owners_tree.insert("", tk.END, values=values)

    def _build_and_draw(self, orgnr: str) -> None:
        """Create OrgChartModel for the given orgnr and draw it on canvas."""
        # Read depth preferences from UI. 0 means unlimited (None)
        up_val = self.depth_up_var.get() if hasattr(self, "depth_up_var") else 2
        down_val = self.depth_down_var.get() if hasattr(self, "depth_down_var") else 2
        max_up = None if up_val == 0 else up_val
        max_down = None if down_val == 0 else down_val
        # Read minimum percentage threshold from UI; treat as percent value (0-100)
        try:
            val = self.min_pct_var.get() if hasattr(self, "min_pct_var") else 0.0
            min_pct = float(val)
        except Exception:
            min_pct = 0.0
        # Reset the canvas zoom to default before drawing a new graph.  This
        # prevents edges from misaligning after repeated zoom operations
        # when switching companies.  Only reset the zoom factor; the canvas
        # will be cleared and redrawn below.
        if hasattr(self.canvas, "_zoom"):
            self.canvas._zoom = 1.0
        # Create model with threshold; percentages below this value are filtered
        self.model = OrgChartModel(
            self.conn,
            root_orgnr=orgnr,
            max_up=max_up,
            max_down=max_down,
            min_pct=min_pct,
        )
        try:
            self.model.build_graph()
        except Exception as exc:
            messagebox.showerror("Feil", f"Kunne ikke bygge graf: {exc}")
            return
        # Attach model to canvas so edges update when dragging
        self.canvas.set_model(self.model)
        # Draw graph
        self.canvas.draw_graph(self.model)
        # Compute scroll region based on positions and node sizes
        # Determine bounding box of all nodes
        if self.model.nodes:
            xs = [n.x for n in self.model.nodes.values()]
            ys = [n.y for n in self.model.nodes.values()]
            width = self.canvas.NODE_WIDTH
            height = self.canvas.NODE_HEIGHT
            x_min = min(xs) - width
            x_max = max(xs) + width
            y_min = min(ys) - height
            y_max = max(ys) + height
            self.canvas.configure(scrollregion=(x_min, y_min, x_max, y_max))

    def _on_node_click(self, node: Node) -> None:
        """Update details panel with information about clicked node."""
        # Map Node attributes to detail fields
        details_map = {
            "Navn": node.name,
            "ID": node.id,
            "Aksjeklasse": node.share_class or "",
            "Landkode": node.country or "",
            "Postnr/sted": node.zip_place or "",
            "Antall aksjer": (f"{int(node.shares_owner_num):,}".replace(",", " ") if node.shares_owner_num is not None else ""),
            "Antall aksjer selskap": (f"{int(node.shares_company_num):,}".replace(",", " ") if node.shares_company_num is not None else ""),
            "Eierandel %": (f"{node.ownership_pct:.2f}" if node.ownership_pct is not None else ""),
        }
        for key, label in self.detail_labels.items():
            val = details_map.get(key, "")
            label.config(text=f"{key}: {val}")

    def _on_node_double_click(self, node: Node) -> None:
        """
        Called when a node in the canvas is double-clicked.  This
        selects the double-clicked company or person as the new root
        of the graph by rebuilding the model and owners table for its ID.

        If the node represents a company, its ID is used directly as
        the organisation number.  If it's a person node (without orgnr),
        we cannot build an ownership graph, so we ignore the double-click.
        """
        # If the node has an ID and is a company, rebuild graph for that ID
        if node and node.id and node.is_company:
            # Set the current_orgnr to the node's id
            self.current_orgnr = node.id
            # Load owners of this new company in the owners table
            self._load_owners(node.id)
            # Build and draw the graph for the new root
            self._build_and_draw(node.id)

    def _on_min_pct_change(self) -> None:
        """Callback when the minimum ownership percentage spinner changes."""
        # If a company is currently selected, rebuild the graph with the new threshold
        if self.current_orgnr:
            # Debounce: avoid too many rapid redraws; schedule redraw after idle
            self.after(200, lambda: self._build_and_draw(self.current_orgnr))

    def _refresh_graph(self) -> None:
        """Reset zoom and redraw the current graph.

        This method resets the canvas zoom back to its default (1.0),
        recenters the scrollbars, and rebuilds the graph for the
        currently selected organisation number.  Useful when the
        layout has become misaligned after extensive zooming and
        panning.
        """
        # Only proceed if a company is selected
        if not self.current_orgnr:
            return
        # Reset zoom on the canvas
        try:
            # Bring zoom back to 1x by scaling all items inversely
            if self.canvas._zoom != 1.0:
                factor = 1.0 / self.canvas._zoom
                self.canvas.scale("all", 0, 0, factor, factor)
                self.canvas._zoom = 1.0
            # Reset scroll position to top-left
            self.canvas.xview_moveto(0)
            self.canvas.yview_moveto(0)
        except Exception:
            pass
        # Redraw graph from model (rebuild not strictly necessary, but
        # ensures lines and positions are recalculated).  We call
        # `_build_and_draw` with the current orgnr so that the graph is
        # reconstructed using existing depth and threshold settings.
        self._build_and_draw(self.current_orgnr)

    # ------------------------------------------------------------------
    # Left panel toggle
    # ------------------------------------------------------------------
    def _toggle_left_panel(self) -> None:
        """
        Collapse or expand the left panel.  When collapsed, the left
        pane containing the search and owners table is removed from
        the PanedWindow so that the canvas occupies all available
        space.  When expanded, the pane is reinserted.
        The toggle button's arrow changes direction accordingly.
        """
        # If currently expanded, remove left_frame
        if not getattr(self, "left_collapsed", False):
            try:
                self.content.forget(self.left_frame)
            except Exception:
                pass
            self.left_collapsed = True
            # Change button arrow to right-pointing triangle
            self.toggle_btn.config(text="▶")  # ▶
        else:
            # Insert left_frame back into panedwindow as first pane
            try:
                # Ensure left_frame isn't already in content
                self.content.add(self.left_frame, weight=1)
                # Optionally reorder panes: PanedWindow adds at end; we want left
                # so we may remove and re-add right_frame to reorder
            except Exception:
                pass
            self.left_collapsed = False
            # Change button arrow to left-pointing triangle
            self.toggle_btn.config(text="◀")  # ◀

    # ------------------------------------------------------------------
    # Public entry point
    # ------------------------------------------------------------------
    def run(self) -> None:
        """Run the Tkinter main event loop."""
        self.mainloop()


if __name__ == "__main__":
    app = OrgChartApp()
    app.run()
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
import db
from org_model import OrgChartModel, Node
from org_view import OrgChartCanvas


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
        self.canvas = OrgChartCanvas(right_frame, on_node_click=self._on_node_click)
        # Put canvas in a scrollable frame to allow panning
        canvas_scroll_y = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        canvas_scroll_x = ttk.Scrollbar(right_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=canvas_scroll_y.set, xscrollcommand=canvas_scroll_x.set)
        # Configure scroll region later when drawing
        self.canvas.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        canvas_scroll_y.pack(fill=tk.Y, side=tk.RIGHT)
        canvas_scroll_x.pack(fill=tk.X, side=tk.BOTTOM)

        # Details panel below canvas
        details_frame = ttk.Labelframe(right_frame, text="Detaljer", padding=6)
        details_frame.pack(fill=tk.X, expand=False, pady=(4, 0))
        # Labels for details (we'll update text when node clicked)
        self.detail_labels = {}
        for field in ["Navn", "ID", "Aksjeklasse", "Landkode", "Postnr/sted", "Antall aksjer", "Antall aksjer selskap", "Eierandel %"]:
            lbl = ttk.Label(details_frame, text=f"{field}: ")
            lbl.pack(anchor="w")
            self.detail_labels[field] = lbl

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
        # Default depth values; could be configurable
        max_up = 2
        max_down = 2
        min_pct = 0.0
        self.model = OrgChartModel(self.conn, root_orgnr=orgnr, max_up=max_up, max_down=max_down, min_pct=min_pct)
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

    # ------------------------------------------------------------------
    # Public entry point
    # ------------------------------------------------------------------
    def run(self) -> None:
        """Run the Tkinter main event loop."""
        self.mainloop()


if __name__ == "__main__":
    app = OrgChartApp()
    app.run()
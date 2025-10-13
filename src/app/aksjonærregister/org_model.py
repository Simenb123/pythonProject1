"""
org_model.py - A model to build shareholder ownership graph from a DuckDB-backed
aksjonærregister.  The model constructs a simple directed graph of nodes
(companies or persons) and edges (ownership relationships) for a specified
root company.  It supports building the graph up to a configurable number
of levels upstream (owners) and downstream (subsidiaries), and filtering
edges below a minimum ownership percentage.

This module is meant to be used together with the interactive Tkinter
visualisation defined in ``org_view.py``.  It relies on the existing
``db.py`` module for data access, so it must live in the same package
(``aksjonærregister``) so that relative imports work.

Example usage::

    from .org_model import OrgChartModel
    from .db import open_conn
    conn = open_conn()
    model = OrgChartModel(conn, root_orgnr="916928092", max_up=2, max_down=1)
    model.build_graph()
    print(model.nodes)
    print(model.edges)

The ``build_graph`` method populates ``model.nodes`` and ``model.edges``.

"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Set, Tuple, Optional

import re

from . import db  # type: ignore


@dataclass
class Node:
    """Represents a company or person in the ownership graph."""
    id: str
    name: str
    # If ``is_company`` is True, this node represents a legal entity with
    # a nine-digit Norwegian organisation number.  Otherwise it is a person
    # identified by birth year, municipality, etc.
    is_company: bool
    # Optional extra fields for detail display
    share_class: Optional[str] = None
    country: Optional[str] = None
    zip_place: Optional[str] = None
    shares_owner_num: Optional[float] = None
    shares_company_num: Optional[float] = None
    ownership_pct: Optional[float] = None
    # Coordinates for layout (to be set by view)
    x: float = 0.0
    y: float = 0.0


@dataclass
class Edge:
    """Represents a directed ownership relationship between two nodes."""
    owner_id: str
    company_id: str
    share_class: Optional[str]
    ownership_pct: Optional[float]


class OrgChartModel:
    """
    Build and store an ownership graph for a company.

    ``OrgChartModel`` constructs a graph of nodes and edges starting from
    a root company organisation number.  It uses functions from ``db.py``
    to query owners (upstream) and children (downstream) from the DuckDB
    database.  It keeps track of visited companies to avoid cycles.
    """

    def __init__(self,
                 conn: db.duckdb.DuckDBPyConnection,
                 root_orgnr: str,
                 max_up: Optional[int] = 2,
                 max_down: Optional[int] = 2,
                 min_pct: float = 0.0) -> None:
        self.conn = conn
        self.root_orgnr = root_orgnr
        # ``max_up`` and ``max_down`` can be ``None`` to indicate unlimited
        # traversal.  Otherwise they indicate the maximum number of levels
        # to traverse in the respective direction (0 means no traversal).
        self.max_up = max_up
        self.max_down = max_down
        self.min_pct = min_pct
        # Mapping from node id to Node object
        self.nodes: Dict[str, Node] = {}
        # List of edges in the graph
        self.edges: List[Edge] = []
        # Track which company orgnrs we have already processed to avoid
        # infinite recursion (for both up and down directions)
        self._visited_up: Set[str] = set()
        self._visited_down: Set[str] = set()

    def build_graph(self) -> None:
        """
        Build the internal graph structures (nodes and edges) for the
        configured root company.  Clears any existing nodes/edges and
        traverses up to ``max_up`` levels of owners and ``max_down`` levels
        of subsidiaries.  Only edges whose percentage is greater than or
        equal to ``min_pct`` are included.
        """
        self.nodes.clear()
        self.edges.clear()
        self._visited_up.clear()
        self._visited_down.clear()

        # Add root company node
        root = self._get_or_create_node(self.root_orgnr, None, is_company=True)
        # If the root node has no descriptive name (name==id), attempt to look up its
        # company name from the database.  The shareholders table stores the
        # company name per row, so we fetch the first non-null entry.  If
        # multiple rows exist for the company, they will all have the same
        # company_name for a given orgnr.  Should the query fail or return
        # no result, we leave the name as the id.
        try:
            row = self.conn.execute(
                "SELECT company_name FROM shareholders WHERE company_orgnr = ? AND company_name IS NOT NULL ORDER BY company_name LIMIT 1",
                [self.root_orgnr],
            ).fetchone()
            if row and row[0]:
                root_node = self.nodes[self.root_orgnr]
                # Only update if the current name is identical to the id (i.e. not already
                # set from a previous query).  This avoids clobbering user-edited
                # names.
                if root_node.name == root_node.id:
                    root_node.name = row[0]
        except Exception:
            # silently ignore lookup errors and keep default name
            pass
        # Recursively build owners and children
        self._add_owners(self.root_orgnr, level=0)
        self._add_children(self.root_orgnr, level=0)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _is_company(self, id_str: str) -> bool:
        """Determine if an identifier represents a company (9 digits)."""
        return bool(re.fullmatch(r"\d{9}", id_str))

    def _get_or_create_node(self, node_id: Optional[str], name: Optional[str], is_company: Optional[bool] = None,
                             **kwargs) -> Node:
        """
        Get an existing node by id or create a new one.  If ``node_id`` is
        blank or None, a temporary identifier will be generated.

        The ``is_company`` flag is optional; if not provided, it will be
        derived from whether ``node_id`` appears to be a 9-digit number.
        Additional keyword arguments can include share_class, country,
        zip_place, shares_owner_num, shares_company_num and ownership_pct.
        """
        if not node_id:
            # Generate a unique temporary ID for anonymous owners
            node_id = f"anonymous_{len(self.nodes)}"
            derived_is_company = False
        else:
            derived_is_company = self._is_company(node_id)

        if is_company is None:
            is_company = derived_is_company

        if node_id not in self.nodes:
            self.nodes[node_id] = Node(
                id=node_id,
                name=name or node_id,
                is_company=is_company,
                share_class=kwargs.get("share_class"),
                country=kwargs.get("country"),
                zip_place=kwargs.get("zip_place"),
                shares_owner_num=kwargs.get("shares_owner_num"),
                shares_company_num=kwargs.get("shares_company_num"),
                ownership_pct=kwargs.get("ownership_pct"),
            )
        else:
            # Update details if not already set
            node = self.nodes[node_id]
            for k in ("name", "share_class", "country", "zip_place", "shares_owner_num",
                      "shares_company_num", "ownership_pct"):
                v = kwargs.get(k)
                if v is not None and getattr(node, k) is None:
                    setattr(node, k, v)
            # Name might be missing initially (owner_name may be None). Update if provided
            if name and node.name == node.id:
                node.name = name
        return self.nodes[node_id]

    def _add_edge(self, owner_id: str, company_id: str, share_class: Optional[str], ownership_pct: Optional[float]) -> None:
        """Add an ownership edge if it does not already exist."""
        for e in self.edges:
            if e.owner_id == owner_id and e.company_id == company_id:
                return  # already exists
        self.edges.append(Edge(owner_id, company_id, share_class, ownership_pct))

    def _add_owners(self, company_orgnr: str, level: int) -> None:
        """
        Recursively add owners (upstream) for ``company_orgnr``.  The ``level``
        argument tracks how many levels up we have gone.  Traversal stops
        when ``level >= max_up`` or the company has already been visited.
        """
        # If max_up is specified, stop when we've reached that depth.  A
        # value of None means unlimited traversal upwards.
        if self.max_up is not None and level >= self.max_up:
            return
        if company_orgnr in self._visited_up:
            return
        self._visited_up.add(company_orgnr)

        # Query owners aggregated per owner
        owner_rows = db.get_owners_agg_owner(self.conn, company_orgnr)
        for row in owner_rows:
            (owner_orgnr, owner_name, share_class, owner_country, owner_zip_place,
             shares_owner_num, shares_company_num, ownership_pct) = row
            # Skip if below threshold
            if ownership_pct is not None and ownership_pct < self.min_pct:
                continue
            # Create owner node
            owner_node = self._get_or_create_node(
                owner_orgnr or None,
                owner_name,
                None,
                share_class=share_class,
                country=owner_country,
                zip_place=owner_zip_place,
                shares_owner_num=shares_owner_num,
                shares_company_num=shares_company_num,
                ownership_pct=ownership_pct,
            )
            # Add edge: owner -> company
            self._add_edge(owner_node.id, company_orgnr, share_class, ownership_pct)
            # Recurse further up if this owner is a company
            if owner_node.is_company and owner_node.id != company_orgnr:
                self._add_owners(owner_node.id, level+1)

    def _add_children(self, owner_orgnr: str, level: int) -> None:
        """
        Recursively add children (downstream) for ``owner_orgnr``.  The
        ``level`` argument tracks how many levels down we have gone.  Traversal
        stops when ``level >= max_down`` or the company has already been
        visited (in the down-direction).
        """
        # If max_down is specified, stop when we've reached that depth.  A
        # value of None means unlimited traversal downwards.
        if self.max_down is not None and level >= self.max_down:
            return
        if owner_orgnr in self._visited_down:
            return
        self._visited_down.add(owner_orgnr)

        # Query children aggregated per company
        child_rows = db.get_children_agg_company(self.conn, owner_orgnr)
        for row in child_rows:
            (child_orgnr, child_name, shares_owner_num, shares_company_num, ownership_pct) = row
            # Skip if below threshold
            if ownership_pct is not None and ownership_pct < self.min_pct:
                continue
            # Create child node
            child_node = self._get_or_create_node(
                child_orgnr or None,
                child_name,
                None,
                shares_owner_num=shares_owner_num,
                shares_company_num=shares_company_num,
                ownership_pct=ownership_pct,
            )
            # Add edge: owner -> child
            self._add_edge(owner_orgnr, child_node.id, None, ownership_pct)
            # Recurse further down if this child is a company
            if child_node.is_company and child_node.id != owner_orgnr:
                self._add_children(child_node.id, level+1)

    # ------------------------------------------------------------------
    # Detail retrieval
    # ------------------------------------------------------------------
    def get_node_details(self, node_id: str) -> Optional[Node]:
        """Return the Node object for ``node_id``, if it exists."""
        return self.nodes.get(node_id)

    def get_owners_of(self, node_id: str) -> List[Tuple[Node, Edge]]:
        """
        Return a list of (owner Node, Edge) pairs where the owner owns
        ``node_id``.  Useful for displaying owners of a selected company.
        """
        result: List[Tuple[Node, Edge]] = []
        for edge in self.edges:
            if edge.company_id == node_id:
                owner = self.nodes.get(edge.owner_id)
                if owner:
                    result.append((owner, edge))
        return result

    def get_children_of(self, node_id: str) -> List[Tuple[Node, Edge]]:
        """
        Return a list of (child Node, Edge) pairs where ``node_id`` owns
        the child.  Useful for displaying subsidiaries of a selected company.
        """
        result: List[Tuple[Node, Edge]] = []
        for edge in self.edges:
            if edge.owner_id == node_id:
                child = self.nodes.get(edge.company_id)
                if child:
                    result.append((child, edge))
        return result
from __future__ import annotations
"""
CLI for aksjonaerregister.

Kommandoer:
  build  – bygg/oppdater DB fra CSV
  search – søk etter selskap
  graph  – generer orgkart (png/html) for selskap
  diag   – (valgfri) diagnostikk av CSV/innlesing – vises bare hvis db.py eksporterer diagnose_csv
"""
import argparse
from typing import Optional

from . import settings as S
from .db import ensure_db, open_conn, search_companies
from .graph import render_graph

# Prøv å hente diagnose-funksjonen hvis den finnes i db.py
try:
    from .db import diagnose_csv as _diagnose_csv  # type: ignore[attr-defined]
    _HAS_DIAG = True
except Exception:
    _HAS_DIAG = False


def cmd_build(args: argparse.Namespace) -> None:
    csv = args.csv or S.CSV_PATH
    ensure_db(csv, S.DB_PATH, delimiter=(args.delimiter or S.DELIMITER), column_map=S.COLUMN_MAP, force=True)
    print(f"Bygd DB: {S.DB_PATH} fra {csv}")


def cmd_search(args: argparse.Namespace) -> None:
    con = open_conn()
    try:
        rows = search_companies(con, args.term, args.by, args.limit)
        for orgnr, name in rows:
            print(f"{orgnr}\t{name}")
    finally:
        con.close()


def cmd_graph(args: argparse.Namespace) -> None:
    con = open_conn()
    try:
        out = render_graph(
            con,
            args.orgnr,
            args.name or args.orgnr,
            mode=args.mode,
            max_up=args.max_up,
            max_down=args.max_down,
        )
        if out:
            print(out)
        else:
            print("Kunne ikke generere graf (mangler Graphviz eller feil ved rendering).")
    finally:
        con.close()


def cmd_diag(args: argparse.Namespace) -> None:
    if not _HAS_DIAG:
        print(
            "Diagnostikk ikke tilgjengelig: db.py eksporterer ikke diagnose_csv().\n"
            "Oppdater db.py til en versjon med diagnose_csv, eller kjør uten 'diag'."
        )
        return
    csv = args.csv or S.CSV_PATH
    delim = args.delimiter or S.DELIMITER
    print(_diagnose_csv(csv, delim, S.COLUMN_MAP))  # type: ignore[misc]


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="aksjonaerregister", description="CLI for aksjonaerregister")
    sub = p.add_subparsers(dest="cmd", required=True)

    # build
    p_build = sub.add_parser("build", help="Bygg DB fra CSV")
    p_build.add_argument("--csv", help="Sti til CSV")
    p_build.add_argument("--delimiter", help='Delimiter (f.eks ";" eller ",")')
    p_build.set_defaults(func=cmd_build)

    # search
    p_search = sub.add_parser("search", help="Søk etter selskap")
    p_search.add_argument("term", help="Søkestreng")
    p_search.add_argument("--by", choices=["navn", "orgnr"], default="navn")
    p_search.add_argument("--limit", type=int, default=50)
    p_search.set_defaults(func=cmd_search)

    # graph
    p_graph = sub.add_parser("graph", help="Generer orgkart for selskap")
    p_graph.add_argument("--orgnr", required=True)
    p_graph.add_argument("--name", default=None)
    p_graph.add_argument("--mode", choices=["up", "down", "both"], default="both")
    p_graph.add_argument("--max-up", type=int, default=S.MAX_DEPTH_UP)
    p_graph.add_argument("--max-down", type=int, default=S.MAX_DEPTH_DOWN)
    p_graph.set_defaults(func=cmd_graph)

    # diag (tilgjengelig uansett; funksjonen sier ifra hvis db.py ikke har diagnose_csv)
    p_diag = sub.add_parser("diag", help="Diagnostikk av CSV → DB (hvis db.py støtter det)")
    p_diag.add_argument("--csv", help="Sti til CSV")
    p_diag.add_argument("--delimiter", help='Delimiter (f.eks ";" eller ",")')
    p_diag.set_defaults(func=cmd_diag)

    return p


def main(argv: Optional[list[str]] = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)
    args.func(args)


if __name__ == "__main__":  # pragma: no cover
    main()


"""
run_orgchart.py - Entry point to launch the interactive org chart app.

This small script simply instantiates and runs the ``OrgChartApp``
defined in ``org_controller.py``.  Use this script to start the
Tkinter-based interface for exploring shareholder ownership graphs.
"""

# Import OrgChartApp.  Use relative import if this file is part of a
# package; otherwise fall back to absolute import.  This allows running
# via ``python -m app.aksjonærregister.run_orgchart`` or directly.
try:
    from .org_controller import OrgChartApp  # type: ignore[import-not-found]
except ImportError:
    from org_controller import OrgChartApp  # type: ignore[import-not-found]


def main() -> None:
    """Entry point for launching the OrgChartApp with optional arguments.

    This function parses command‑line flags and passes them to
    ``OrgChartApp``.  You can specify an initial organisation number
    via ``--orgnr``, enable editing mode via ``--editable``, and set
    a JSON file for storing node positions via ``--layout``.  When
    editing is disabled (default), the org chart opens in read‑only
    mode, suitable for the central aksjonærregister lookup.  When
    editing is enabled, node positions can be dragged and saved to
    the layout file on exit.
    """
    import argparse

    parser = argparse.ArgumentParser(description="Start the interactive org chart GUI.")
    parser.add_argument("--orgnr", dest="orgnr", type=str, default=None,
                        help="Organisation number to preselect as root")
    parser.add_argument("--editable", action="store_true",
                        help="Enable editing (dragging/multi‑select). Default is read‑only.")
    parser.add_argument("--layout", dest="layout", type=str, default=None,
                        help="Path to JSON file for saving/loading node positions")

    args = parser.parse_args()
    # Instantiate the app with supplied options.  editable defaults to False
    # unless explicitly passed.  layout_path may be None.
    app = OrgChartApp(
        editable=args.editable,
        root_orgnr=args.orgnr,
        layout_path=args.layout,
    )
    app.run()


if __name__ == "__main__":
    main()
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
    app = OrgChartApp()
    app.run()


if __name__ == "__main__":
    main()
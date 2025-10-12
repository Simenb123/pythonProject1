"""
run_orgchart.py - Entry point to launch the interactive org chart app.

This small script simply instantiates and runs the ``OrgChartApp``
defined in ``org_controller.py``.  Use this script to start the
Tkinter-based interface for exploring shareholder ownership graphs.
"""

from org_controller import OrgChartApp


def main() -> None:
    app = OrgChartApp()
    app.run()


if __name__ == "__main__":
    main()
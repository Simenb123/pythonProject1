"""Widgets package for GUI components.

This package contains reusable widgets used in the client portal, such as
tables and version panels.  The modules here were adapted from the original
`gui/widgets` package and are provided to avoid import errors when the
`KildefilerView` and other views refer to `VersionsPanel`.

Exports:
    DataTable -- simple tabular display for pandas DataFrame objects
    VersionsPanel -- UI component for managing versions of source files
"""

from .data_table import DataTable  # noqa: F401
from .versions_panel import VersionsPanel  # noqa: F401
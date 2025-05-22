# Bilagsverktøy

Simple utilities for selecting a client and working with voucher files.

## Requirements

The scripts require Python 3 with GUI support plus a few libraries:

- `pandas`
- `openpyxl`
- `tkinter` (bundled with most Python installations)

Install packages with:

```bash
pip install pandas openpyxl
```

## Usage

Start `klientvelger.py` to choose a client folder and launch the voucher GUI:

```bash
python klientvelger.py
```

You can also run `bilag_gui_tk.py` directly by giving a path to your voucher
file:

```bash
python bilag_gui_tk.py /path/to/vouchers.xlsx
```

The GUI shows live statistics under the input fields. These numbers refresh
every 300 ms while you adjust the selection.

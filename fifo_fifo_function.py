"""
Simple FIFO helper for Visindi share‑transaction list
====================================================
Assumptions
-----------
* The raw data – after Power Query – has **one row per owner per transaction**
  with these columns (case‑sensitive):

    - "Dato"          : datetime64[ns]  – transaction date
    - "Selskap"       : str             – company / instrument
    - "Eier"          : str             – owner (buyer / seller)
    - "Netto"         : int             – +shares for buy, –shares for sale
    - "Pris pr aksje" : float | int     – price per share (optional)

* Positive *Netto* adds to the position, negative *Netto* reduces it.
* When position goes negative (short) the logic still works, but costBasis
  becomes 0 until there are earlier lots to offset – tweak if you need a
  different behaviour.

Outputs
-------
* A *realised* DataFrame with realisedQty, realisedCost, realisedProceeds,
  pnl for each sale row (other rows are NaN)
* An *unrealised* DataFrame with the closing lot‑stack for each owner ×
  company after the whole dataframe is processed.
"""
from __future__ import annotations

from collections import deque
from dataclasses import dataclass
from typing import Deque, Dict, List, Tuple

import pandas as pd

Lot = Tuple[int, float]  # (quantity, unit_cost)


def _consume_fifo(stack: Deque[Lot], qty_to_sell: int) -> Tuple[int, float]:
    """Remove *qty_to_sell* from *stack* (FIFO) and return (cost, matched_qty).
    If the stack is exhausted before *qty_to_sell* is filled, only the
    available quantity is matched and cost is computed for that amount.
    """
    cost_accum = 0.0
    matched = 0

    while qty_to_sell and stack:
        lot_qty, lot_cost = stack[0]
        take = min(lot_qty, qty_to_sell)
        cost_accum += take * lot_cost
        matched += take
        qty_to_sell -= take
        if take == lot_qty:
            stack.popleft()
        else:
            stack[0] = (lot_qty - take, lot_cost)

    return matched, cost_accum


@dataclass
class FIFOState:
    stack: Deque[Lot]
    realised_rows: List[Dict]


def fifo(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Apply FIFO per (Selskap, Eier).

    Parameters
    ----------
    df : pandas.DataFrame
        Must contain the columns described in the module docstring.

    Returns
    -------
    realised, closing_stacks : Tuple[pd.DataFrame, pd.DataFrame]
        *realised*  – row‑wise realised result for sales (qty<0)
        *closing*   – each owner×company remaining stack after last row
    """
    required = {"Dato", "Selskap", "Eier", "Netto"}
    if not required.issubset(df.columns):
        missing = required - set(df.columns)
        raise ValueError(f"Missing required columns: {missing}")

    # Sort by date just in case
    df_sorted = df.sort_values("Dato").reset_index(drop=True)

    # state per (company, owner)
    book: Dict[Tuple[str, str], FIFOState] = {}

    realised_records: List[Dict] = []

    for idx, row in df_sorted.iterrows():
        key = (row["Selskap"], row["Eier"])
        qty: int = int(row["Netto"])
        price = float(row.get("Pris pr aksje", 0))

        # ensure state exists
        state = book.setdefault(key, FIFOState(deque(), []))

        if qty > 0:
            # purchase – push onto stack
            state.stack.append((qty, price))
        elif qty < 0:
            # sale – pull from stack FIFO
            qty_to_sell = -qty
            matched, cost = _consume_fifo(state.stack, qty_to_sell)
            proceeds = matched * price
            realised_records.append(
                {
                    "Dato": row["Dato"],
                    "Selskap": key[0],
                    "Eier": key[1],
                    "QtySold": matched,
                    "Proceeds": proceeds,
                    "Cost": cost,
                    "PnL": proceeds - cost,
                }
            )

    # Build closing stack dataframe
    closing = []
    for (company, owner), state in book.items():
        for lot_qty, lot_cost in state.stack:
            closing.append({
                "Selskap": company,
                "Eier": owner,
                "Qty": lot_qty,
                "CostPerShare": lot_cost,
            })

    realised_df = pd.DataFrame(realised_records)
    closing_df = pd.DataFrame(closing)

    return realised_df, closing_df


if __name__ == "__main__":
    # Minimal example (remove in production)
    data = {
        "Dato": pd.to_datetime(["2024-01-01", "2024-02-01", "2024-03-01"]),
        "Selskap": ["Visindi AS"] * 3,
        "Eier": ["Atle Ronglan AS"] * 3,
        "Netto": [1000, -600, -200],
        "Pris pr aksje": [50, 60, 65],
    }
    test_df = pd.DataFrame(data)
    realised, closing = fifo(test_df)
    print("Realised:\n", realised)
    print("Closing stack:\n", closing)

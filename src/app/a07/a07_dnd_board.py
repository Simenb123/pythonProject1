# a07_dnd_board.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox

try:
    from tkinterdnd2 import TkinterDnD, DND_TEXT
    DND_OK = True
except Exception:
    # Kjører uten DnD (GUI virker, men drag&drop deaktivert)
    TkinterDnD = None
    DND_TEXT = None
    DND_OK = False


class DnDBoard(ttk.Frame):
    """
    En enkel DnD-brett:
      - Venstre: A07-mål (koder + bundles), viser sum, diff og status
      - Høyre: Tilgjengelige GL-konti
    public API:
      - set_targets(list[(id, label, amount, diff_status)])  diff_status: "ok"/"warn"/"none"
      - set_accounts(list[(konto, label)])
      - clear_mapping_for(target_id) / clear_account(konto)
      - bind("<<DropAccountOnTarget>>", handler)  # event.data = f"{konto}|{target_id}"
      - bind("<<UnmapTarget>>", handler)          # event.data = target_id
      - set_hide_mapped(bool)
    """

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        # venstre: targets
        left = ttk.Frame(self)
        right = ttk.Frame(self)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        ttk.Label(left, text="A07-koder / Bundles").pack(anchor="w")
        self.target_list = tk.Listbox(left, height=18, activestyle="none")
        self.target_scroll = ttk.Scrollbar(left, command=self.target_list.yview)
        self.target_list.config(yscrollcommand=self.target_scroll.set)
        self.target_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.target_scroll.pack(side=tk.LEFT, fill=tk.Y)

        ttk.Label(right, text="GL-konti (tilgjengelige)").pack(anchor="w")
        self.acc_list = tk.Listbox(right, height=18, activestyle="none")
        self.acc_scroll = ttk.Scrollbar(right, command=self.acc_list.yview)
        self.acc_list.config(yscrollcommand=self.acc_scroll.set)
        self.acc_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.acc_scroll.pack(side=tk.LEFT, fill=tk.Y)

        # data
        self._targets: list[tuple[str, str, float, str]] = []  # (id, label, amount, status)
        self._accounts: list[tuple[int, str]] = []
        self._hide_mapped = tk.BooleanVar(value=True)

        # høyreklikk på targets -> "fjern mapping"
        self.target_menu = tk.Menu(self, tearoff=False)
        self.target_menu.add_command(label="Fjern mapping", command=self._popup_unmap)

        self.target_list.bind("<Button-3>", self._show_menu)
        self.target_list.bind("<ButtonRelease-3>", lambda e: "break")

        # DnD
        if DND_OK:
            # Kilder: konto-lista
            self.acc_list.drop_target_register()
            self.acc_list.drag_source_register(1, DND_TEXT)
            self.acc_list.dnd_bind("<<DragInitCmd>>", self._on_drag_init)
            # Mål: target-lista
            self.target_list.drop_target_register(DND_TEXT)
            self.target_list.dnd_bind("<<Drop>>", self._on_drop)

    # ---------------- Public API -----------------

    def set_hide_mapped(self, hide: bool):
        self._hide_mapped.set(bool(hide))
        self._refresh_lists()

    def set_targets(self, items: list[tuple[str, str, float, str]]):
        """
        items: [(target_id, label, diff_amount, status)]
        status in {"ok","warn","none"}
        """
        self._targets = list(items)
        self._refresh_lists()

    def set_accounts(self, items: list[tuple[int, str]]):
        self._accounts = list(items)
        self._refresh_lists()

    def mark_target_status(self, target_id: str, status: str):
        for i, (tid, lbl, diff, st) in enumerate(self._targets):
            if tid == target_id:
                self._targets[i] = (tid, lbl, diff, status)
                break
        self._refresh_lists()

    # ---------------- Internal -------------------

    def _refresh_lists(self):
        self.target_list.delete(0, tk.END)
        for tid, label, diff, status in self._targets:
            txt = label
            if status == "ok":
                txt += "  ✓"
            elif status == "warn":
                txt += "  (diff)"
            self.target_list.insert(tk.END, txt)
            idx = self.target_list.size() - 1
            if status == "ok":
                self.target_list.itemconfig(idx, background="#d4edda")  # grønnlig
            elif status == "warn":
                self.target_list.itemconfig(idx, background="#fff3cd")  # gul
            else:
                self.target_list.itemconfig(idx, background="#f8f9fa")  # lys grå

        self.acc_list.delete(0, tk.END)
        for konto, label in self._accounts:
            self.acc_list.insert(tk.END, label)

    def _on_drag_init(self, event):
        try:
            idxs = list(self.acc_list.curselection())
            if not idxs:
                return
            idx = idxs[0]
            konto = self._accounts[idx][0]
            return (DND_TEXT, tk.dnd.Icon(""), f"{konto}", 0)
        except Exception:
            return

    def _on_drop(self, event):
        try:
            data = str(event.data).strip()
            if not data:
                return
            # finn valgt target
            idxs = list(self.target_list.curselection())
            if not idxs:
                return
            tid = self._targets[idxs[0]][0]
            payload = f"{data}|{tid}"
            self.event_generate("<<DropAccountOnTarget>>", data=payload)
        except Exception as e:
            messagebox.showerror("DnD-feil", str(e))

    def _show_menu(self, event):
        try:
            idx = self.target_list.nearest(event.y)
            self.target_list.selection_clear(0, tk.END)
            self.target_list.selection_set(idx)
            self.target_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.target_menu.grab_release()

    def _popup_unmap(self):
        idxs = self.target_list.curselection()
        if not idxs:
            return
        tid = self._targets[idxs[0]][0]
        self.event_generate("<<UnmapTarget>>", data=tid)

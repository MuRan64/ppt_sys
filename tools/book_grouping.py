"""Tk dialogs and validation for user-defined PPT chapter ranges."""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Any


class UserCancelled(Exception):
    """Raised when the user intentionally cancels a grouping workflow."""


def suggested_ranges(
    units: list[dict[str, Any]],
    max_per_group: int,
    *,
    preserve_sections: bool = True,
) -> list[dict[str, Any]]:
    """Make editable default ranges while preserving each source section."""

    if max_per_group < 1:
        raise ValueError("每组篇目数必须大于零。")
    ranges: list[dict[str, Any]] = []
    sections: list[tuple[str, list[dict[str, Any]]]] = []
    for unit in units:
        volume = str(unit["volume"])
        if not sections or sections[-1][0] != volume:
            sections.append((volume, []))
        sections[-1][1].append(unit)

    if not preserve_sections:
        sections = [("全部篇目", units)]

    for volume, members in sections:
        size = max_per_group
        if volume == "真题速递":
            size = 2
        for start in range(0, len(members), size):
            selected = members[start : start + size]
            ranges.append(
                {
                    "volume": volume,
                    "unit_keys": [(unit["volume"], unit["num"]) for unit in selected],
                    "start": selected[0]["num"],
                    "end": selected[-1]["num"],
                }
            )
    return ranges


class BookGroupingDialog(tk.Toplevel):
    """Show detected units and let users repartition them into contiguous ranges."""

    def __init__(
        self,
        parent: tk.Misc,
        *,
        book_name: str,
        units: list[dict[str, Any]],
        max_per_group: int,
        allow_cross_sections: bool = False,
        initial_ranges: list[dict[str, Any]] | None = None,
    ) -> None:
        if not units:
            raise ValueError("未识别到可分组的篇目。请核对所选书籍与正文文件是否对应。")
        keys = [(str(unit["volume"]), int(unit["num"])) for unit in units]
        if len(keys) != len(set(keys)):
            raise ValueError("扫描结果存在同章节重复篇目编号，请先核对正文结构，避免答案错配。")
        super().__init__(parent)
        self.title(f"{book_name} - PPT 分组")
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        width = min(1080, max(760, screen_w - 100))
        height = min(800, max(600, screen_h - 140))
        self.geometry(f"{width}x{height}")
        self.minsize(min(760, screen_w), min(560, screen_h))
        self.transient(parent)
        self.grab_set()
        self.units = units
        self.result: list[dict[str, Any]] | None = None
        self.allow_cross_sections = allow_cross_sections
        self.groups = initial_ranges or suggested_ranges(
            units, max_per_group, preserve_sections=not allow_cross_sections
        )
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        outer = ttk.Frame(self, padding=14)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=3)
        outer.rowconfigure(2, weight=2)
        ttk.Label(
            outer,
            text=(
                f"已从转换后的 MD 识别到 {len(units)} 个篇目/练习。"
                + ("按扫描顺序连续分组；下方是当前建议，可自行调整。" if allow_cross_sections else "按专题内的篇目范围分组；下方是当前建议，可自行调整。")
            ),
            wraplength=900,
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        unit_frame = ttk.LabelFrame(outer, text="扫描结果", padding=8)
        unit_frame.grid(row=1, column=0, sticky="nsew")
        unit_columns = ("index", "volume", "num", "title")
        self.unit_tree = ttk.Treeview(unit_frame, columns=unit_columns, show="headings", selectmode="extended", height=15)
        for key, title, width, anchor in (
            ("index", "序号", 55, "center"),
            ("volume", "章节/专题", 220, "w"),
            ("num", "篇目/练", 90, "center"),
            ("title", "识别到的标题", 480, "w"),
        ):
            self.unit_tree.heading(key, text=title)
            self.unit_tree.column(key, width=width, anchor=anchor, stretch=key == "title")
        unit_scroll = ttk.Scrollbar(unit_frame, orient="vertical", command=self.unit_tree.yview)
        self.unit_tree.configure(yscrollcommand=unit_scroll.set)
        self.unit_tree.pack(side="left", fill="both", expand=True)
        unit_scroll.pack(side="right", fill="y")
        for index, unit in enumerate(units):
            self.unit_tree.insert(
                "",
                "end",
                iid=str(index),
                values=(index + 1, unit["volume"], unit["num"], unit["title"]),
            )

        group_frame = ttk.LabelFrame(outer, text="PPT 分组", padding=8)
        group_frame.grid(row=2, column=0, sticky="nsew", pady=(10, 0))
        group_columns = ("index", "volume", "range", "titles")
        self.group_tree = ttk.Treeview(group_frame, columns=group_columns, show="headings", selectmode="extended", height=7)
        for key, title, width, anchor in (
            ("index", "PPT", 55, "center"),
            ("volume", "章节/专题", 220, "w"),
            ("range", "篇目范围", 140, "center"),
            ("titles", "包含篇目", 480, "w"),
        ):
            self.group_tree.heading(key, text=title)
            self.group_tree.column(key, width=width, anchor=anchor, stretch=key == "titles")
        group_scroll = ttk.Scrollbar(group_frame, orient="vertical", command=self.group_tree.yview)
        self.group_tree.configure(yscrollcommand=group_scroll.set)
        self.group_tree.pack(side="left", fill="both", expand=True)
        group_scroll.pack(side="right", fill="y")

        actions = ttk.Frame(outer)
        actions.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions, text="按所选篇目新建分组", command=self._add_selection, width=23).pack(side="left")
        ttk.Button(actions, text="删除选中分组", command=self._remove_groups, width=16).pack(side="left", padx=(8, 0))
        ttk.Label(
            outer,
            text="选择连续篇目；所有扫描到的篇目都必须且只能归入一个分组。",
            foreground="#555",
        ).grid(row=4, column=0, sticky="w", pady=(6, 0))

        footer = ttk.Frame(outer)
        footer.grid(row=5, column=0, sticky="ew", pady=(12, 0))
        ttk.Button(footer, text="取消", command=self._cancel, width=12).pack(side="right", padx=(8, 0))
        ttk.Button(footer, text="按分组生成 PPT", command=self._accept, width=18).pack(side="right")
        self._refresh_groups()

    def _unit_key(self, unit: dict[str, Any]) -> tuple[str, int]:
        return str(unit["volume"]), int(unit["num"])

    def _selected_units(self) -> list[dict[str, Any]]:
        selected = [self.units[int(iid)] for iid in self.unit_tree.selection()]
        return sorted(selected, key=lambda unit: self.units.index(unit))

    def _add_selection(self) -> None:
        selected = self._selected_units()
        if len(selected) < 1:
            messagebox.showwarning("未选择篇目", "请先在扫描结果中选择一个连续范围。", parent=self)
            return
        volumes = {str(unit["volume"]) for unit in selected}
        if len(volumes) != 1 and not self.allow_cross_sections:
            messagebox.showwarning("范围跨章节", "一个 PPT 分组不能跨越不同章节/专题，请分开设置。", parent=self)
            return
        volume = str(selected[0]["volume"]) if len(volumes) == 1 else "多个章节"
        ordered_units = self.units if self.allow_cross_sections else [unit for unit in self.units if str(unit["volume"]) == volume]
        positions = [ordered_units.index(unit) for unit in selected]
        if positions != list(range(min(positions), max(positions) + 1)):
            messagebox.showwarning("篇目不连续", "请选择同一专题内连续的篇目。", parent=self)
            return
        self.groups.append(
            {
                "volume": volume,
                "unit_keys": [self._unit_key(unit) for unit in selected],
                "start": int(selected[0]["num"]),
                "end": int(selected[-1]["num"]),
            }
        )
        unit_positions = {self._unit_key(unit): index for index, unit in enumerate(self.units)}
        self.groups.sort(key=lambda group: unit_positions[tuple(group["unit_keys"][0])])
        self._refresh_groups()

    def _remove_groups(self) -> None:
        indexes = sorted((int(iid) for iid in self.group_tree.selection()), reverse=True)
        for index in indexes:
            del self.groups[index]
        self._refresh_groups()

    def _refresh_groups(self) -> None:
        for child in self.group_tree.get_children():
            self.group_tree.delete(child)
        for index, group in enumerate(self.groups):
            members = [
                next(unit for unit in self.units if self._unit_key(unit) == tuple(key))
                for key in group["unit_keys"]
            ]
            title_text = "、".join(str(unit["title"]) for unit in members)
            self.group_tree.insert(
                "",
                "end",
                iid=str(index),
                values=(index + 1, group["volume"], f"{group['start']}—{group['end']}", title_text),
            )

    def _accept(self) -> None:
        if not self.groups:
            messagebox.showerror("尚无分组", "请至少设置一个 PPT 分组。", parent=self)
            return
        grouped_keys = [tuple(key) for group in self.groups for key in group["unit_keys"]]
        expected_keys = [self._unit_key(unit) for unit in self.units]
        if grouped_keys != expected_keys:
            messagebox.showerror(
                "分组不完整",
                "当前分组必须按原目录顺序覆盖全部篇目，且不能重复或遗漏。请删除相关分组后重新划分。",
                parent=self,
            )
            return
        self.result = [dict(group) for group in self.groups]
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()

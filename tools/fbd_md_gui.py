#!/usr/bin/env python3
"""Small desktop app for converting selected FBD files into styled Markdown.

Current workflow:
1. Select a body FBD file.
2. Select an answer FBD file.
3. Optionally select supplement/booklet FBD files when a book needs them.
4. Select a PPTX template file.
5. Click convert. The app writes Markdown files into a timestamped local folder.
6. When a supported subject/book adapter is selected, continue generating PPTX.
"""

from __future__ import annotations  

import json
import queue
import sys
import threading
import traceback
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from tkinter import END, LEFT, RIGHT, BOTH, X, Y, filedialog, messagebox, ttk
import tkinter as tk


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = Path(getattr(sys, "_MEIPASS", SCRIPT_DIR.parent))
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from app_resources import open_directory, user_data_directory  # noqa: E402
from template_library import templates_for_book, validate_pptx_template  # noqa: E402


def default_output_base() -> Path:
    # A bundled executable may live in Program Files or a temporary extraction
    # directory. Generated user files must remain writable and survive exit.
    if getattr(sys, "frozen", False):
        return user_data_directory()
    return PROJECT_DIR


DEFAULT_OUTPUT_ROOT = default_output_base() / "output" / "gui_md"
DEFAULT_PPT_OUTPUT_ROOT = default_output_base() / "output" / "gui_ppt"

# Allow importing the converter when this script is launched from any cwd.
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))
TESTS_DIR = PROJECT_DIR / "tests"
if str(TESTS_DIR) not in sys.path:
    sys.path.insert(0, str(TESTS_DIR))

from fbd_to_md import (  # noqa: E402
    convert_fbd_to_markdown,
    load_tif_map,
    normalize_tif_key,
    read_fbd,
    save_tif_map,
    tif_markdown_tag,
)

from test_md_to_ppt import generate_wenyan_pptx  # noqa: E402
from test_md_to_ppt import scan_wenyan_group_units  # noqa: E402
from test_sanbu_md_to_ppt import generate_sanbu_pptx, scan_sanbu_group_units  # noqa: E402
from test_english_md_to_ppt import generate_english_book_pptx  # noqa: E402
from test_yilun_xinfangan_md_to_ppt import generate_yilun_pptx  # noqa: E402
from test_yilun_xinfangan_md_to_ppt import scan_yilun_group_units  # noqa: E402
from book_grouping import BookGroupingDialog, UserCancelled  # noqa: E402
from book_adapters.bibei_gushiwen import generate_bibei_book_pptx_from_fbd  # noqa: E402
from book_adapters.wenyan_yuedu_jiafen import generate_wenyan_yuedu_jiafen_pptx_from_fbd  # noqa: E402
from book_adapters.zuhe_jiafen import generate_zuhe_jiafen_pptx_from_fbd  # noqa: E402
from book_adapters.zhuji_100 import generate_zhuji_100_pptx_from_fbd  # noqa: E402


SUBJECT_BOOKS = {
    "语文": [
        "文言梳理",
        "三步进阶",
        "必背古诗文",
        "高考组合加分练",
        "文言文阅读加分练",
        "筑基臻选100练",
        "高考语文大一轮复习新方案",
    ],
    "英语": ["词汇语境记忆"],
}

EXTRA_INPUT_REQUIREMENTS = {
    ("语文", "必背古诗文"): {"extra": "必背古诗文册子.fbd", "pocket": "必背口袋本.fbd"},
    ("英语", "词汇语境记忆"): {"extra": "词汇语境记忆册子.fbd"},
}


class TifPromptDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk, *, tif_name: str, image_path: Path | None, command: str) -> None:
        super().__init__(parent)
        self.title("确认 FBD 图片文字")
        self.resizable(False, False)
        self.result: dict[str, str] = {"action": "ignore"}
        self.preview_image = None

        container = ttk.Frame(self, padding=16)
        container.pack(fill=BOTH, expand=True)

        ttk.Label(
            container,
            text="FBD 中引用了一个 tif 图片，请确认它是否需要保留到 MD。",
            font=("Microsoft YaHei UI", 11, "bold"),
        ).pack(anchor="w")
        ttk.Label(container, text=f"文件：{tif_name}", foreground="#555").pack(anchor="w", pady=(6, 0))
        if image_path:
            ttk.Label(container, text=str(image_path), foreground="#777", wraplength=520).pack(anchor="w")

        preview_frame = ttk.Frame(container, padding=(0, 12, 0, 8))
        preview_frame.pack(fill=X)
        self._add_preview(preview_frame, image_path)

        ttk.Label(
            container,
            text="如果这是正文里需要保留的单字/文字图片，点击“保留到 MD”；如果是背景、线条、装饰图，点击“忽略此图”。",
            wraplength=520,
        ).pack(anchor="w", pady=(0, 10))

        command_preview = command.replace("\n", " ")
        if len(command_preview) > 90:
            command_preview = command_preview[:90] + "..."
        ttk.Label(container, text=f"原始引用：{command_preview}", foreground="#777", wraplength=520).pack(anchor="w")

        actions = ttk.Frame(container)
        actions.pack(fill=X, pady=(14, 0))
        ttk.Button(actions, text="保留到 MD", command=self._keep).pack(side=LEFT)
        ttk.Button(actions, text="忽略此图", command=self._ignore).pack(side=LEFT, padx=(10, 0))

        self.bind("<Return>", lambda _event: self._keep())
        self.bind("<Escape>", lambda _event: self._ignore())
        self.protocol("WM_DELETE_WINDOW", self._ignore)
        self.transient(parent)
        self.grab_set()
        self.update_idletasks()
        x = parent.winfo_rootx() + max(0, (parent.winfo_width() - self.winfo_width()) // 2)
        y = parent.winfo_rooty() + max(0, (parent.winfo_height() - self.winfo_height()) // 2)
        self.geometry(f"+{x}+{y}")

    def _add_preview(self, parent: ttk.Frame, image_path: Path | None) -> None:
        if not image_path or not image_path.exists():
            ttk.Label(parent, text="未找到对应 tif 文件，无法预览。", foreground="#b00020").pack(anchor="w")
            return
        try:
            from PIL import Image, ImageTk

            image = Image.open(image_path)
            image.thumbnail((260, 180))
            self.preview_image = ImageTk.PhotoImage(image)
            ttk.Label(parent, image=self.preview_image).pack(anchor="w")
        except Exception as exc:
            ttk.Label(
                parent,
                text=f"无法预览 tif：{exc}\n可安装 pillow 后显示图片预览。",
                foreground="#b00020",
                wraplength=520,
            ).pack(anchor="w")

    def _keep(self) -> None:
        self.result = {"action": "keep"}
        self.destroy()

    def _ignore(self) -> None:
        self.result = {"action": "ignore"}
        self.destroy()


class TemplateSelector(ttk.Frame):
    """Select a book's bundled template or retain its independent custom file."""

    def __init__(self, parent: tk.Misc, path_variable: tk.StringVar) -> None:
        super().__init__(parent)
        self.path_variable = path_variable
        self.mode_variable = tk.StringVar(value="内置模板")
        self.choice_variable = tk.StringVar()
        self.options: dict[str, Path] = {}
        self._context: tuple[str, str, str] | None = None
        self._saved: dict[tuple[str, str, str], tuple[str, str, str]] = {}
        self._custom_path = ""
        self._active_mode = "内置模板"

        self.mode_combo = ttk.Combobox(self, textvariable=self.mode_variable, values=("内置模板", "自选模板"), state="readonly", width=9)
        self.mode_combo.grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.mode_combo.bind("<<ComboboxSelected>>", self._mode_changed)
        self.builtin_combo = ttk.Combobox(self, textvariable=self.choice_variable, state="readonly", width=24)
        self.builtin_combo.grid(row=0, column=1, sticky="ew")
        self.builtin_combo.bind("<<ComboboxSelected>>", lambda _event: self._apply_mode())
        self.custom_entry = ttk.Entry(self, textvariable=self.path_variable)
        self.custom_entry.grid(row=0, column=1, sticky="ew")
        self.custom_entry.grid_remove()
        self.browse_button = ttk.Button(self, text="选择文件...", command=self._browse)
        self.browse_button.grid(row=0, column=2, padx=(8, 0))
        self.columnconfigure(1, weight=1)

    def set_book(self, subject: str, book: str, slot: str = "primary") -> None:
        context = (subject, book, slot)
        if context == self._context:
            return
        if self._context is not None:
            if self._active_mode == "自选模板":
                self._custom_path = self.path_variable.get()
            self._saved[self._context] = (self.mode_variable.get(), self.choice_variable.get(), self._custom_path)
        try:
            self.options = {template.label: template.path for template in templates_for_book(subject, book, slot)}
        except (OSError, ValueError):
            self.options = {}
        self._context = context
        mode, choice, self._custom_path = self._saved.get(context, ("内置模板", "", ""))
        if not self.options:
            mode = "自选模板"
        self.mode_variable.set(mode)
        self.choice_variable.set(choice if choice in self.options else next(iter(self.options), ""))
        self.builtin_combo.configure(values=list(self.options))
        self._apply_mode()

    def _mode_changed(self, _event=None) -> None:
        if self._active_mode == "自选模板":
            self._custom_path = self.path_variable.get()
        self._apply_mode()

    def _apply_mode(self) -> None:
        self._active_mode = self.mode_variable.get()
        if self._active_mode == "内置模板":
            self.custom_entry.grid_remove()
            self.builtin_combo.grid()
            path = self.options.get(self.choice_variable.get())
            self.path_variable.set(str(path) if path else "")
        else:
            self.builtin_combo.grid_remove()
            self.custom_entry.grid()
            self.path_variable.set(self._custom_path)

    def _browse(self) -> None:
        current = self.path_variable.get().strip()
        initial_dir = str(Path(current).parent) if current else None
        selected = filedialog.askopenfilename(parent=self, title="选择 PPTX 模板", filetypes=[("PowerPoint files", "*.pptx")], initialdir=initial_dir)
        if selected:
            self._custom_path = selected
            self.mode_variable.set("自选模板")
            self._apply_mode()


class FbdMdApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("FBD 转 Markdown 工具")
        self.geometry("920x620")
        self.minsize(820, 560)

        self.body_fbd_var = tk.StringVar()
        self.answer_fbd_var = tk.StringVar()
        self.extra_fbd_var = tk.StringVar()
        self.pocket_fbd_var = tk.StringVar()
        self.template_pptx_var = tk.StringVar()
        self.module2_template_var = tk.StringVar()
        self.output_root_var = tk.StringVar(value=str(DEFAULT_OUTPUT_ROOT))
        self.subject_var = tk.StringVar(value="语文")
        self.book_var = tk.StringVar(value="文言梳理")
        self.status_var = tk.StringVar(value="请选择正文 FBD、答案 FBD、PPTX 模板和书籍适配。")

        self.message_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self._input_row_widgets: dict[int, list[tk.Widget]] = {}

        self._build_ui()
        self._update_book_options()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(120, self._poll_queue)

    def _on_close(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("任务运行中", "当前任务尚未结束，请等待任务完成；如果正在设置分组，可以在分组窗口点击取消。", parent=self)
            return
        self.destroy()

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=18)
        root.pack(fill=BOTH, expand=True)

        title = ttk.Label(root, text="FBD 转 Markdown 工具", font=("Microsoft YaHei UI", 18, "bold"))
        title.pack(anchor="w")

        subtitle = ttk.Label(
            root,
            text="选择已适配书目时会先生成 Markdown，再继续按模板生成 PPTX。",
            foreground="#555",
        )
        subtitle.pack(anchor="w", pady=(4, 16))

        form = ttk.Frame(root)
        form.pack(fill=X)

        self._adapter_row(form, row=0)
        self._file_row(
            form,
            row=1,
            label="正文 .fbd",
            variable=self.body_fbd_var,
            filetypes=[("FBD files", "*.fbd"), ("All files", "*.*")],
        )
        self._file_row(
            form,
            row=2,
            label="答案 .fbd",
            variable=self.answer_fbd_var,
            filetypes=[("FBD files", "*.fbd"), ("All files", "*.*")],
        )
        self._file_row(
            form,
            row=3,
            label="补充/册子 .fbd",
            variable=self.extra_fbd_var,
            filetypes=[("FBD files", "*.fbd"), ("All files", "*.*")],
        )
        self._file_row(
            form,
            row=4,
            label="口袋本 .fbd",
            variable=self.pocket_fbd_var,
            filetypes=[("FBD files", "*.fbd"), ("All files", "*.*")],
        )
        self.primary_template_label = ttk.Label(form, text="PPTX 模板", width=14)
        self.primary_template_label.grid(row=5, column=0, sticky="w", pady=7)
        self.template_selector = TemplateSelector(form, self.template_pptx_var)
        self.template_selector.grid(row=5, column=1, columnspan=2, sticky="ew", padx=(8, 0), pady=7)
        module2_label = ttk.Label(form, text="模块二模板", width=14)
        module2_label.grid(row=6, column=0, sticky="w", pady=7)
        self.module2_template_selector = TemplateSelector(form, self.module2_template_var)
        self.module2_template_selector.grid(row=6, column=1, columnspan=2, sticky="ew", padx=(8, 0), pady=7)
        self._input_row_widgets[6] = [module2_label, self.module2_template_selector]
        self._folder_row(form, row=7, label="MD 输出根目录", variable=self.output_root_var)

        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(root)
        actions.pack(fill=X, pady=(18, 10))

        self.convert_button = ttk.Button(actions, text="开始生成", command=self._start_convert)
        self.convert_button.pack(side=LEFT)

        open_button = ttk.Button(actions, text="打开输出根目录", command=self._open_output_root)
        open_button.pack(side=LEFT, padx=(10, 0))

        self.progress = ttk.Progressbar(actions, mode="indeterminate", length=220)
        self.progress.pack(side=RIGHT)

        status = ttk.Label(root, textvariable=self.status_var, foreground="#6b126b")
        status.pack(anchor="w", pady=(4, 8))

        log_frame = ttk.LabelFrame(root, text="运行日志", padding=8)
        log_frame.pack(fill=BOTH, expand=True)

        self.log_text = tk.Text(log_frame, height=16, wrap="word", font=("Consolas", 10))
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)

        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scroll.pack(side=RIGHT, fill=Y)
        self.log_text.configure(yscrollcommand=scroll.set)

    def _file_row(
        self,
        parent: ttk.Frame,
        *,
        row: int,
        label: str,
        variable: tk.StringVar,
        filetypes: list[tuple[str, str]],
    ) -> None:
        label_widget = ttk.Label(parent, text=label, width=14)
        label_widget.grid(row=row, column=0, sticky="w", pady=7)
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", padx=(8, 8), pady=7)
        button = ttk.Button(
            parent,
            text="选择...",
            command=lambda: self._choose_file(variable, filetypes),
        )
        button.grid(row=row, column=2, sticky="e", pady=7)
        row_widgets: list[tk.Widget] = [label_widget, entry, button]
        self._input_row_widgets[row] = row_widgets
        if row == 3:
            self.extra_row_label = label_widget

    def _folder_row(self, parent: ttk.Frame, *, row: int, label: str, variable: tk.StringVar) -> None:
        ttk.Label(parent, text=label, width=14).grid(row=row, column=0, sticky="w", pady=7)
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", padx=(8, 8), pady=7)
        button = ttk.Button(parent, text="选择...", command=lambda: self._choose_folder(variable))
        button.grid(row=row, column=2, sticky="e", pady=7)

    def _adapter_row(self, parent: ttk.Frame, *, row: int) -> None:
        ttk.Label(parent, text="PPT 适配", width=14).grid(row=row, column=0, sticky="w", pady=7)
        controls = ttk.Frame(parent)
        controls.grid(row=row, column=1, sticky="ew", padx=(8, 8), pady=7)
        controls.columnconfigure(1, weight=1)

        ttk.Label(controls, text="科目").grid(row=0, column=0, sticky="w")
        self.subject_combo = ttk.Combobox(
            controls,
            textvariable=self.subject_var,
            values=list(SUBJECT_BOOKS),
            state="readonly",
            width=12,
        )
        self.subject_combo.grid(row=0, column=1, sticky="w", padx=(8, 18))
        self.subject_combo.bind("<<ComboboxSelected>>", lambda _event: self._update_book_options())

        ttk.Label(controls, text="书名").grid(row=0, column=2, sticky="w")
        self.book_combo = ttk.Combobox(
            controls,
            textvariable=self.book_var,
            values=SUBJECT_BOOKS["语文"],
            state="readonly",
            width=18,
        )
        self.book_combo.grid(row=0, column=3, sticky="w", padx=(8, 0))
        self.book_combo.bind("<<ComboboxSelected>>", lambda _event: self._update_required_inputs())
        ttk.Label(parent, text="选择已适配书目时会在 MD 后继续生成 PPT。", foreground="#555").grid(
            row=row, column=2, sticky="e", pady=7
        )

    def _update_book_options(self) -> None:
        books = SUBJECT_BOOKS.get(self.subject_var.get(), [])
        self.book_combo.configure(values=books)
        if books:
            if self.book_var.get() not in books:
                self.book_var.set(books[0])
            self.book_combo.configure(state="readonly")
        else:
            self.book_var.set("")
            self.book_combo.configure(state="disabled")
        self._update_required_inputs()

    def _update_required_inputs(self) -> None:
        requirements = EXTRA_INPUT_REQUIREMENTS.get((self.subject_var.get(), self.book_var.get()), {})
        extra_label = requirements.get("extra")
        pocket_label = requirements.get("pocket")
        previous = getattr(self, "_current_extra_requirements", {})
        if previous.get("extra") != extra_label:
            self.extra_fbd_var.set("")
        if previous.get("pocket") != pocket_label:
            self.pocket_fbd_var.set("")
        self._current_extra_requirements = requirements
        self._set_input_row_visible(3, bool(extra_label))
        self._set_input_row_visible(4, bool(pocket_label))
        if hasattr(self, "extra_row_label"):
            self.extra_row_label.configure(text=extra_label or "补充/册子 .fbd")
        is_zhuji = (self.subject_var.get(), self.book_var.get()) == ("语文", "筑基臻选100练")
        self.primary_template_label.configure(text="模块一模板" if is_zhuji else "PPTX 模板")
        self._set_input_row_visible(6, is_zhuji)
        self.template_selector.set_book(self.subject_var.get(), self.book_var.get())
        self.module2_template_selector.set_book(self.subject_var.get(), self.book_var.get(), "secondary")

    def _set_input_row_visible(self, row: int, visible: bool) -> None:
        for widget in self._input_row_widgets.get(row, []):
            if visible:
                widget.grid()
            else:
                widget.grid_remove()

    def _choose_file(self, variable: tk.StringVar, filetypes: list[tuple[str, str]]) -> None:
        initial_dir = None
        current = variable.get().strip()
        if current:
            path = Path(current)
            initial_dir = str(path.parent if path.parent.exists() else PROJECT_DIR)
        selected = filedialog.askopenfilename(title="选择文件", filetypes=filetypes, initialdir=initial_dir)
        if selected:
            variable.set(selected)

    def _choose_folder(self, variable: tk.StringVar) -> None:
        current = variable.get().strip()
        initial_dir = current if current and Path(current).exists() else str(PROJECT_DIR)
        selected = filedialog.askdirectory(title="选择输出目录", initialdir=initial_dir)
        if selected:
            variable.set(selected)

    def _open_output_root(self) -> None:
        try:
            output_root = Path(self.output_root_var.get().strip() or DEFAULT_OUTPUT_ROOT)
            output_root.mkdir(parents=True, exist_ok=True)
            # Windows shell open. The project is Windows-based, and this keeps
            # the GUI dependency-free.
            open_directory(output_root)
        except Exception as exc:
            messagebox.showerror("无法打开目录", str(exc))

    def _log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(END, f"[{timestamp}] {message}\n")
        self.log_text.see(END)

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.convert_button.configure(state=state)
        # Freeze the form while the worker uses its captured input paths.
        form = self.subject_combo.master.master

        def set_controls(parent: tk.Misc) -> None:
            for widget in parent.winfo_children():
                if isinstance(widget, ttk.Combobox):
                    widget.configure(state="disabled" if busy else "readonly")
                elif isinstance(widget, (ttk.Entry, ttk.Button)):
                    widget.configure(state=state)
                set_controls(widget)

        set_controls(form)
        if busy:
            self.progress.start(10)
        else:
            self.progress.stop()

    def _validate_inputs(self) -> tuple[Path, Path, Path | None, Path | None, Path, Path]:
        if self.book_var.get() not in SUBJECT_BOOKS.get(self.subject_var.get(), []):
            raise ValueError("请先选择已适配的科目和书籍。")
        body_text = self.body_fbd_var.get().strip()
        answer_text = self.answer_fbd_var.get().strip()
        template_text = self.template_pptx_var.get().strip()
        if not body_text:
            raise ValueError("请先选择：正文 .fbd")
        if not answer_text:
            raise ValueError("请先选择：答案 .fbd")
        if not template_text:
            raise ValueError("请先选择：PPTX 模板")

        body_fbd = Path(body_text)
        answer_fbd = Path(answer_text)
        requirements = EXTRA_INPUT_REQUIREMENTS.get((self.subject_var.get(), self.book_var.get()), {})
        extra_text = self.extra_fbd_var.get().strip() if requirements.get("extra") else ""
        extra_fbd = Path(extra_text) if extra_text else None
        pocket_text = self.pocket_fbd_var.get().strip() if requirements.get("pocket") else ""
        pocket_fbd = Path(pocket_text) if pocket_text else None
        template_pptx = Path(template_text)
        output_root = Path(self.output_root_var.get().strip() or DEFAULT_OUTPUT_ROOT)

        required = [
            ("正文 .fbd", body_fbd, ".fbd"),
            ("答案 .fbd", answer_fbd, ".fbd"),
            ("PPTX 模板", template_pptx, ".pptx"),
        ]
        for label, path, suffix in required:
            if not str(path):
                raise ValueError(f"请先选择：{label}")
            if not path.is_file():
                raise FileNotFoundError(f"{label} 不存在：{path}")
            if path.suffix.lower() != suffix:
                raise ValueError(f"{label} 扩展名应为 {suffix}：{path}")

        if requirements.get("extra") and not extra_fbd:
            raise ValueError(f"《{self.book_var.get()}》还需要选择：{requirements['extra']}")
        if requirements.get("pocket") and not pocket_fbd:
            raise ValueError(f"《{self.book_var.get()}》还需要选择：{requirements['pocket']}")
        if extra_fbd:
            if not extra_fbd.is_file():
                raise FileNotFoundError(f"补充/册子 .fbd 不存在：{extra_fbd}")
            if extra_fbd.suffix.lower() != ".fbd":
                raise ValueError(f"补充/册子 .fbd 扩展名应为 .fbd：{extra_fbd}")
        if pocket_fbd:
            if not pocket_fbd.is_file():
                raise FileNotFoundError(f"口袋本 .fbd 不存在：{pocket_fbd}")
            if pocket_fbd.suffix.lower() != ".fbd":
                raise ValueError(f"口袋本 .fbd 扩展名应为 .fbd：{pocket_fbd}")

        fbd_paths = [path for path in (body_fbd, answer_fbd, extra_fbd, pocket_fbd) if path is not None]
        if len({str(path.resolve()).casefold() for path in fbd_paths}) != len(fbd_paths):
            raise ValueError("正文、答案、册子和口袋本不能选择同一个文件，请分别选择对应文件。")
        for path in fbd_paths:
            if path.stat().st_size == 0:
                raise ValueError(f"FBD 文件为空：{path}")
        validate_pptx_template(template_pptx, minimum_slides=7 if self.subject_var.get() == "英语" else 6)
        if self.book_var.get() == "筑基臻选100练":
            module2_text = self.module2_template_var.get().strip()
            if not module2_text:
                raise ValueError("请先选择：模块二模板")
            validate_pptx_template(Path(module2_text), label="模块二模板")

        output_root.mkdir(parents=True, exist_ok=True)
        return body_fbd, answer_fbd, extra_fbd, pocket_fbd, template_pptx, output_root

    def _selected_adapter(self) -> str | None:
        subject = self.subject_var.get()
        book = self.book_var.get()
        if subject == "语文" and book in {
            "文言梳理",
            "三步进阶",
            "必背古诗文",
            "高考组合加分练",
            "文言文阅读加分练",
            "筑基臻选100练",
            "高考语文大一轮复习新方案",
        }:
            return f"{subject}/{book}"
        if subject == "英语" and book == "词汇语境记忆":
            return f"{subject}/{book}"
        return None

    def _find_tif_image(self, tif_name: str, search_dirs: list[Path]) -> Path | None:
        wanted = normalize_tif_key(tif_name)
        for directory in search_dirs:
            if not directory.exists():
                continue
            direct = directory / tif_name
            if direct.exists():
                return direct
            for path in directory.rglob("*.tif"):
                if normalize_tif_key(path.name) == wanted:
                    return path
            for path in directory.rglob("*.tiff"):
                if normalize_tif_key(path.name) == wanted:
                    return path
        return None

    def _prompt_tif_from_main_thread(self, request: dict[str, object]) -> None:
        event = request.get("event")
        try:
            image_path_text = str(request.get("image_path") or "")
            image_path = Path(image_path_text) if image_path_text else None
            dialog = TifPromptDialog(
                self,
                tif_name=str(request.get("tif_name") or ""),
                image_path=image_path,
                command=str(request.get("command") or ""),
            )
            self.wait_window(dialog)
            request["response"] = dialog.result
        except Exception as exc:
            request["error"] = exc
        finally:
            if isinstance(event, threading.Event):
                event.set()

    def _prompt_grouping_from_main_thread(self, request: dict[str, object]) -> None:
        event = request.get("event")
        try:
            units = request.get("units")
            if not isinstance(units, list):
                request["response"] = None
            else:
                dialog = BookGroupingDialog(
                    self,
                    book_name=str(request.get("book") or ""),
                    units=units,
                    max_per_group=int(request.get("max_per_group") or 3),
                    allow_cross_sections=bool(request.get("allow_cross_sections", False)),
                    initial_ranges=request.get("initial_ranges") if isinstance(request.get("initial_ranges"), list) else None,
                )
                self.wait_window(dialog)
                request["response"] = dialog.result
        except Exception as exc:
            request["error"] = exc
        finally:
            if isinstance(event, threading.Event):
                event.set()

    def _select_book_groups(
        self,
        book: str,
        units: list[dict[str, object]],
        *,
        max_per_group: int | None = None,
        allow_cross_sections: bool = False,
        initial_ranges: list[dict[str, object]] | None = None,
    ) -> list[dict[str, object]] | None:
        if not units:
            raise ValueError(f"《{book}》未识别到可分组篇目，请检查所选书名和正文 FBD 是否对应。")
        event = threading.Event()
        request: dict[str, object] = {
            "book": book,
            "units": units,
            "max_per_group": max_per_group or (3 if book == "高考语文组合加分练" else 5),
            "allow_cross_sections": allow_cross_sections,
            "initial_ranges": initial_ranges,
            "event": event,
            "response": None,
        }
        self.message_queue.put(("group_prompt", request))
        event.wait()
        prompt_error = request.get("error")
        if isinstance(prompt_error, Exception):
            raise prompt_error
        response = request.get("response")
        if not isinstance(response, list):
            raise UserCancelled(f"用户取消了《{book}》的分组设置。")
        return response

    def _require_book_groups(self, book: str, units: list[dict[str, object]]) -> list[dict[str, object]]:
        result = self._select_book_groups(book, units)
        if result is None:
            raise UserCancelled(f"用户取消了《{book}》的 PPT 分组设置。")
        return result

    def _start_convert(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            return

        try:
            inputs = self._validate_inputs()
        except Exception as exc:
            messagebox.showerror("输入有误", str(exc))
            return

        self.log_text.delete("1.0", END)
        self._set_busy(True)
        self.status_var.set("正在转换，请稍候...")
        self._log("开始转换任务。")

        subject = self.subject_var.get()
        book = self.book_var.get()
        module2_template = Path(self.module2_template_var.get().strip()) if book == "筑基臻选100练" else None
        self.worker_thread = threading.Thread(target=self._convert_worker, args=(*inputs, subject, book), kwargs={"module2_template": module2_template}, daemon=True)
        self.worker_thread.start()

    def _convert_worker(
        self,
        body_fbd: Path,
        answer_fbd: Path,
        extra_fbd: Path | None,
        pocket_fbd: Path | None,
        template_pptx: Path,
        output_root: Path,
        subject: str,
        book: str,
        module2_template: Path | None = None,
    ) -> None:
        try:
            run_name = datetime.now().strftime("run_%Y%m%d_%H%M%S_%f")
            run_dir = output_root / run_name
            run_dir.mkdir(parents=True, exist_ok=False)

            self.message_queue.put(("log", f"输出目录：{run_dir}"))
            if subject == "语文" and book == "必背古诗文":
                if extra_fbd is None or pocket_fbd is None:
                    raise ValueError("语文《必背古诗文》需要同时选择册子 FBD 和口袋本 FBD。")
                ppt_output_dir = DEFAULT_PPT_OUTPUT_ROOT / run_name / "必背古诗文"
                self.message_queue.put(("log", "已选择适配：语文/必背古诗文，开始转换四个 FBD 并生成主书/小册子 PPT。"))
                result = generate_bibei_book_pptx_from_fbd(
                    body_fbd=body_fbd,
                    answer_fbd=answer_fbd,
                    booklet_fbd=extra_fbd,
                    pocket_fbd=pocket_fbd,
                    template=template_pptx,
                    md_output_dir=run_dir,
                    ppt_output_dir=ppt_output_dir,
                    group_selector=lambda label, units, defaults: self._select_book_groups(
                        f"必背古诗文 - {label}",
                        units,
                        max_per_group=6 if label == "主书" else 8,
                        allow_cross_sections=label == "主书",
                        initial_ranges=defaults,
                    ),
                    progress=lambda message: self.message_queue.put(("log", message)),
                )
                self.message_queue.put(("log", f"PPT 已生成 {len(result.ppt_outputs)} 个，输出目录：{result.ppt_output_dir}"))
                self.message_queue.put(("log", f"适配器清单：{result.manifest_path}"))
                self.message_queue.put(("done", {"md_dir": str(run_dir), "ppt_dir": str(result.ppt_output_dir)}))
                return
            if subject == "语文" and book == "高考组合加分练":
                ppt_output_dir = DEFAULT_PPT_OUTPUT_ROOT / run_name / "高考组合加分练"
                self.message_queue.put(("log", "已选择适配：语文/高考组合加分练，开始转换正文/答案 FBD 并生成 PPT。"))
                result = generate_zuhe_jiafen_pptx_from_fbd(
                    body_fbd=body_fbd,
                    answer_fbd=answer_fbd,
                    template=template_pptx,
                    md_output_dir=run_dir,
                    ppt_output_dir=ppt_output_dir,
                    group_selector=lambda units: self._select_book_groups("高考语文组合加分练", units),
                    progress=lambda message: self.message_queue.put(("log", message)),
                )
                self.message_queue.put(("log", f"PPT 已生成 {len(result.ppt_outputs)} 个，输出目录：{result.ppt_output_dir}"))
                self.message_queue.put(("log", f"适配器清单：{result.manifest_path}"))
                self.message_queue.put(("done", {"md_dir": str(result.md_run), "ppt_dir": str(result.ppt_output_dir)}))
                return
            if subject == "语文" and book == "文言文阅读加分练":
                ppt_output_dir = DEFAULT_PPT_OUTPUT_ROOT / run_name / "文言文阅读加分练"
                self.message_queue.put(("log", "已选择适配：语文/文言文阅读加分练，开始转换正文/答案 FBD 并生成 PPT。"))
                result = generate_wenyan_yuedu_jiafen_pptx_from_fbd(
                    body_fbd=body_fbd,
                    answer_fbd=answer_fbd,
                    template=template_pptx,
                    md_output_dir=run_dir,
                    ppt_output_dir=ppt_output_dir,
                    group_selector=lambda units: self._select_book_groups("文言文阅读加分练", units),
                    progress=lambda message: self.message_queue.put(("log", message)),
                )
                self.message_queue.put(("log", f"PPT 已生成 {len(result.ppt_outputs)} 个，输出目录：{result.ppt_output_dir}"))
                self.message_queue.put(("log", f"适配器清单：{result.manifest_path}"))
                self.message_queue.put(("done", {"md_dir": str(result.md_run), "ppt_dir": str(result.ppt_output_dir)}))
                return
            if subject == "语文" and book == "筑基臻选100练":
                ppt_output_dir = DEFAULT_PPT_OUTPUT_ROOT / run_name / "筑基臻选100练"
                self.message_queue.put(("log", "已选择适配：语文/筑基臻选100练，开始转换正文/答案 FBD 并生成双模板 PPT。"))
                result = generate_zhuji_100_pptx_from_fbd(
                    body_fbd=body_fbd,
                    answer_fbd=answer_fbd,
                    template=template_pptx,
                    module2_template=module2_template,
                    md_output_dir=run_dir,
                    ppt_output_dir=ppt_output_dir,
                    group_selector=lambda units: self._require_book_groups("筑基臻选100练", units),
                    progress=lambda message: self.message_queue.put(("log", message)),
                )
                self.message_queue.put(("log", f"PPT 已生成 {len(result.ppt_outputs)} 个，输出目录：{result.ppt_output_dir}"))
                self.message_queue.put(("log", f"适配器清单：{result.manifest_path}"))
                self.message_queue.put(("done", {"md_dir": str(result.md_run), "ppt_dir": str(result.ppt_output_dir)}))
                return
            tif_map_path = body_fbd.parent / "fbd_tif_map.json"
            try:
                persisted_tif_replacements, persisted_tif_ignores = load_tif_map(tif_map_path)
            except (ValueError, TypeError, AttributeError, OSError) as exc:
                persisted_tif_replacements, persisted_tif_ignores = {}, set()
                self.message_queue.put(("log", f"旧 TIF 映射不可用，本次重新选择：{exc}"))
            session_tif_replacements: dict[str, str] = {}
            session_tif_ignores: set[str] = set()
            tif_search_dirs = [body_fbd.parent, answer_fbd.parent]
            if extra_fbd:
                tif_search_dirs.append(extra_fbd.parent)
            if pocket_fbd:
                tif_search_dirs.append(pocket_fbd.parent)
            tif_search_dirs = list(dict.fromkeys(tif_search_dirs))
            self.message_queue.put(("log", f"tif 映射文件：{tif_map_path}"))
            auto_ignore_tif = subject == "英语"
            if auto_ignore_tif:
                self.message_queue.put(("log", "英语适配将自动忽略 FBD 中的 tif 引用，不弹出确认框。"))
            else:
                self.message_queue.put(("log", "本次运行会逐个确认 FBD 中出现的 tif；同名 tif 只询问一次。"))

            def resolve_tif(tif_name: str, command: str) -> str | None:
                key = normalize_tif_key(tif_name)
                if auto_ignore_tif:
                    session_tif_ignores.add(key)
                    return None
                if key in session_tif_replacements:
                    return session_tif_replacements[key]
                if key in session_tif_ignores:
                    return None

                image_path = self._find_tif_image(tif_name, tif_search_dirs)
                event = threading.Event()
                request: dict[str, object] = {
                    "tif_name": tif_name,
                    "image_path": str(image_path) if image_path else "",
                    "command": command,
                    "event": event,
                    "response": None,
                }
                self.message_queue.put(("tif_prompt", request))
                event.wait()
                prompt_error = request.get("error")
                if isinstance(prompt_error, Exception):
                    raise prompt_error
                response = request.get("response")
                if isinstance(response, dict) and response.get("action") == "keep":
                    value = tif_markdown_tag(tif_name)
                    session_tif_replacements[key] = value
                    session_tif_ignores.discard(key)
                    persisted_tif_replacements[key] = value
                    persisted_tif_ignores.discard(key)
                    persist_tif_choices()
                    self.message_queue.put(("log", f"tif 保留：{tif_name}"))
                    return value
                session_tif_ignores.add(key)
                session_tif_replacements.pop(key, None)
                persisted_tif_ignores.add(key)
                persisted_tif_replacements.pop(key, None)
                persist_tif_choices()
                self.message_queue.put(("log", f"tif 忽略：{tif_name}"))
                return None

            def persist_tif_choices() -> None:
                try:
                    save_tif_map(tif_map_path, persisted_tif_replacements, persisted_tif_ignores)
                except OSError:
                    # Read-only source folders must not prevent PPT generation.
                    save_tif_map(run_dir / "fbd_tif_map.json", persisted_tif_replacements, persisted_tif_ignores)

            outputs = []
            conversion_items = [("正文", body_fbd), ("答案", answer_fbd)]
            if extra_fbd:
                conversion_items.append(("册子", extra_fbd))
            for role, fbd_path in conversion_items:
                self.message_queue.put(("log", f"读取{role} FBD：{fbd_path}"))
                raw = read_fbd(fbd_path, "gb18030")
                markdown, stats = convert_fbd_to_markdown(
                    raw,
                    auto_start=book != "高考语文大一轮复习新方案",
                    tif_replacements=session_tif_replacements,
                    tif_ignores=session_tif_ignores,
                    tif_resolver=resolve_tif,
                )
                md_path = run_dir / f"{role}_{fbd_path.stem}_格式保留.md"
                md_path.write_text(markdown, encoding="utf-8", newline="\n")
                outputs.append(
                    {
                        "role": role,
                        "source": str(fbd_path),
                        "markdown": str(md_path),
                        "stats": asdict(stats),
                    }
                )
                self.message_queue.put(
                    (
                        "log",
                        f"{role} MD 已生成：{md_path.name} "
                        f"(加点 {stats.dots}，横线 {stats.underlines}，"
                        f"波浪线 {stats.waves}，填空线 {stats.blanks}，上标 {stats.superscripts}，"
                        f"tif保留 {stats.tif_kept}，tif忽略 {stats.tif_ignored}，tif未知 {stats.tif_unknown})",
                    )
                )

            # Generators need source paths now, not only after PPT generation.
            (run_dir / "conversion_manifest.json").write_text(
                json.dumps({"outputs": outputs, "tif_source_dirs": [str(path) for path in tif_search_dirs]}, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            adapter = None
            if subject == "语文" and book in {"文言梳理", "三步进阶"}:
                adapter = f"{subject}/{book}"
            elif subject == "语文" and book == "高考语文大一轮复习新方案":
                adapter = f"{subject}/{book}"
            elif subject == "英语" and book == "词汇语境记忆":
                adapter = f"{subject}/{book}"
            ppt_outputs: list[str] = []
            ppt_output_dir: Path | None = None
            if adapter == "语文/文言梳理":
                ppt_output_dir = DEFAULT_PPT_OUTPUT_ROOT / run_name / "文言梳理"
                self.message_queue.put(("log", f"已选择适配：{adapter}，开始由 MD 生成 PPT。"))
                generated = generate_wenyan_pptx(
                    md_run=run_dir,
                    template=template_pptx,
                    output_dir=ppt_output_dir,
                    limit_groups=None,
                    group_ranges=self._require_book_groups("文言梳理", scan_wenyan_group_units(run_dir)),
                    progress=lambda message: self.message_queue.put(("log", message)),
                )
                ppt_outputs = [str(path) for path in generated]
                self.message_queue.put(("log", f"PPT 已生成 {len(ppt_outputs)} 个，输出目录：{ppt_output_dir}"))
            elif adapter == "语文/三步进阶":
                ppt_output_dir = DEFAULT_PPT_OUTPUT_ROOT / run_name / "三步进阶"
                self.message_queue.put(("log", f"已选择适配：{adapter}，开始由 MD 生成 PPT。"))
                generated = generate_sanbu_pptx(
                    md_run=run_dir,
                    template=template_pptx,
                    output_dir=ppt_output_dir,
                    limit_groups=None,
                    group_ranges=self._require_book_groups("三步进阶", scan_sanbu_group_units(run_dir)),
                    progress=lambda message: self.message_queue.put(("log", message)),
                )
                ppt_outputs = [str(path) for path in generated]
                self.message_queue.put(("log", f"PPT 已生成 {len(ppt_outputs)} 个，输出目录：{ppt_output_dir}"))
            elif adapter == "英语/词汇语境记忆":
                ppt_output_dir = DEFAULT_PPT_OUTPUT_ROOT / run_name / "词汇语境记忆"
                self.message_queue.put(("log", f"已选择适配：{adapter}，开始由 MD 生成 PPT。"))
                generated = generate_english_book_pptx(
                    md_run=run_dir,
                    template=template_pptx,
                    output_dir=ppt_output_dir,
                    passage_nums=None,
                    progress=lambda message: self.message_queue.put(("log", message)),
                )
                ppt_outputs = [str(path) for path in generated]
                self.message_queue.put(("log", f"PPT 已生成 {len(ppt_outputs)} 个，输出目录：{ppt_output_dir}"))
            elif adapter == "语文/高考语文大一轮复习新方案":
                ppt_output_dir = DEFAULT_PPT_OUTPUT_ROOT / run_name / "高考语文大一轮复习新方案"
                self.message_queue.put(("log", f"已选择适配：{adapter}，开始由 MD 生成分专题 PPT。"))
                generated = generate_yilun_pptx(
                    md_run=run_dir,
                    template=template_pptx,
                    output_dir=ppt_output_dir,
                    group_ranges=self._require_book_groups("高考语文大一轮复习新方案", scan_yilun_group_units(run_dir)),
                    progress=lambda message: self.message_queue.put(("log", message)),
                )
                ppt_outputs = [str(path) for path in generated]
                self.message_queue.put(("log", f"PPT 已生成 {len(ppt_outputs)} 个，输出目录：{ppt_output_dir}"))
            else:
                self.message_queue.put(("log", "未选择可用 PPT 适配器，本次只生成 Markdown。"))

            manifest = {
                "created_at": datetime.now().isoformat(timespec="seconds"),
                "template_pptx": str(template_pptx),
                "output_dir": str(run_dir),
                "tif_map": str(tif_map_path),
                "subject": subject,
                "book": book,
                "ppt_adapter": adapter,
                "ppt_output_dir": str(ppt_output_dir) if ppt_output_dir else "",
                "ppt_outputs": ppt_outputs,
                "outputs": outputs,
                "note": "选择已适配书目时会继续根据 MD 和 PPTX 模板生成 PPT；其他适配暂只生成 Markdown。",
            }
            manifest_path = run_dir / "conversion_manifest.json"
            manifest_path.write_text(
                json.dumps(manifest, ensure_ascii=False, indent=2),
                encoding="utf-8",
                newline="\n",
            )
            self.message_queue.put(("log", f"清单已生成：{manifest_path.name}"))
            self.message_queue.put(("done", {"md_dir": str(run_dir), "ppt_dir": str(ppt_output_dir) if ppt_output_dir else ""}))
        except UserCancelled as exc:
            self.message_queue.put(("cancelled", str(exc)))
        except Exception:
            self.message_queue.put(("error", traceback.format_exc()))

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self.message_queue.get_nowait()
                if kind == "log":
                    self._log(str(payload))
                elif kind == "tif_prompt":
                    if isinstance(payload, dict):
                        self._prompt_tif_from_main_thread(payload)
                elif kind == "group_prompt":
                    if isinstance(payload, dict):
                        self._prompt_grouping_from_main_thread(payload)
                elif kind == "done":
                    if isinstance(payload, dict):
                        run_dir = Path(str(payload.get("md_dir") or ""))
                        ppt_dir = str(payload.get("ppt_dir") or "")
                    else:
                        run_dir = Path(str(payload))
                        ppt_dir = ""
                    self._set_busy(False)
                    self.status_var.set(f"转换完成：{run_dir}")
                    self._log("任务完成。")
                    message = f"Markdown 已生成到：\n{run_dir}"
                    if ppt_dir:
                        message += f"\n\nPPT 已生成到：\n{ppt_dir}"
                    messagebox.showinfo("转换完成", message)
                elif kind == "cancelled":
                    self._set_busy(False)
                    self.status_var.set("操作已取消。")
                    self._log("用户已取消，本次未继续生成 PPT。")
                    messagebox.showinfo("已取消", "操作已取消，未继续生成 PPT。已转换的 Markdown 保留在本次运行目录中。")
                elif kind == "error":
                    self._set_busy(False)
                    self.status_var.set("转换失败，请查看日志。")
                    self._log(str(payload))
                    messagebox.showerror("转换未完成", "本次转换未完成。请确认输入文件、配套模板及输出目录可用后重试。详细原因已保留在运行日志中。")
        except queue.Empty:
            pass
        self.after(120, self._poll_queue)


def main() -> int:
    app = FbdMdApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

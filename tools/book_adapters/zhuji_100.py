#!/usr/bin/env python3
"""Adapter for 语文《筑基臻选100练》.

The tested layout rules live in ``tests/test_zhuji_100_md_to_ppt.py``.  This
production-facing adapter keeps the GUI orchestration small and book-scoped:
convert the selected source/answer FBD files into the MD run expected by the
generator, then generate the module-one and module-two PPTX groups with their
separate templates.
"""

from __future__ import annotations

import json
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable


TOOLS_DIR = Path(__file__).resolve().parents[1]
PROJECT_DIR = Path(getattr(sys, "_MEIPASS", TOOLS_DIR.parent))
TESTS_DIR = PROJECT_DIR / "tests"
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))
if str(TESTS_DIR) not in sys.path:
    sys.path.insert(0, str(TESTS_DIR))

import test_zhuji_100_md_to_ppt as zhuji_ppt  # noqa: E402
from book_grouping import UserCancelled  # noqa: E402
from template_library import builtin_template_path, validate_pptx_template  # noqa: E402


Progress = Callable[[str], None]
GroupSelector = Callable[[list[dict[str, object]]], list[dict[str, object]] | None]

DEFAULT_MODULE1_TEMPLATE = builtin_template_path("筑基模块一模板.pptx")
DEFAULT_MODULE2_TEMPLATE = builtin_template_path("筑基模块二模板.pptx")


@dataclass(frozen=True)
class Zhuji100GenerationResult:
    """Return value for GUI callers."""

    md_run: Path
    ppt_output_dir: Path
    ppt_outputs: list[Path]
    manifest_path: Path
    module1_template: Path
    module2_template: Path


def _emit(progress: Progress | None, message: str) -> None:
    if progress:
        progress(message)


def _write_json(path: Path, payload: dict[str, object]) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8", newline="\n")


def _first_existing(candidates: list[Path]) -> Path:
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


def _validate_template(path: Path, label: str) -> None:
    validate_pptx_template(path, label=label)


def resolve_zhuji_templates(selected_template: Path, module2_template: Path | None = None) -> tuple[Path, Path]:
    """Resolve the two fixed templates used by this book.

    The GUI has one template picker because earlier books use one template.
    《筑基臻选100练》 has two templates, so the selected file is treated as an
    override for the matching module and the companion template is discovered
    from the project/template directory.
    """

    if module2_template is not None:
        _validate_template(selected_template, "筑基模块一模板")
        _validate_template(module2_template, "筑基模块二模板")
        return selected_template, module2_template

    selected_parent = selected_template.parent
    module1_candidates = [selected_parent / DEFAULT_MODULE1_TEMPLATE.name, DEFAULT_MODULE1_TEMPLATE]
    module2_candidates = [selected_parent / DEFAULT_MODULE2_TEMPLATE.name, DEFAULT_MODULE2_TEMPLATE]

    module1_template = _first_existing(module1_candidates)
    module2_template = _first_existing(module2_candidates)

    name = selected_template.stem
    if "模块一" in name:
        module1_template = selected_template
    elif "模块二" in name:
        module2_template = selected_template
    else:
        module1_template = selected_template

    _validate_template(module1_template, "筑基模块一模板")
    _validate_template(module2_template, "筑基模块二模板")
    return module1_template, module2_template


def generate_zhuji_100_pptx_from_fbd(
    *,
    body_fbd: Path,
    answer_fbd: Path,
    template: Path,
    md_output_dir: Path,
    ppt_output_dir: Path,
    module2_template: Path | None = None,
    group_selector: GroupSelector | None = None,
    progress: Progress | None = None,
) -> Zhuji100GenerationResult:
    """Full production entry point for GUI: FBD pair -> MD run -> PPTX."""

    module1_template, module2_template = resolve_zhuji_templates(template, module2_template)
    _emit(progress, f"筑基臻选100练：模块一模板：{module1_template}")
    _emit(progress, f"筑基臻选100练：模块二模板：{module2_template}")

    md_run = zhuji_ppt.convert_zuhe_fbd_to_md(
        body_fbd=body_fbd,
        answer_fbd=answer_fbd,
        output_root=md_output_dir,
        progress=progress,
    )

    all_units = zhuji_ppt.scan_zhuji_group_units(md_run)
    if not all_units:
        raise ValueError("筑基臻选100练正文未识别到练习，请核对所选书籍及正文文件。")
    module2_volumes = {"一年新题通关", "三年模拟题通关", "创新预测题通关", "组合通关60练"}
    module2_units = [unit for unit in all_units if str(unit["volume"]) in module2_volumes]
    module1_units = [unit for unit in all_units if str(unit["volume"]) not in module2_volumes]
    module1_ranges = None
    module2_ranges = None
    if group_selector:
        if module1_units:
            module1_ranges = group_selector(module1_units)
            if module1_ranges is None:
                raise UserCancelled("用户取消了筑基臻选100练模块一 PPT 分组设置。")
        else:
            module1_ranges = []
        if module2_units:
            module2_ranges = group_selector(module2_units)
            if module2_ranges is None:
                raise UserCancelled("用户取消了筑基臻选100练模块二 PPT 分组设置。")
        else:
            module2_ranges = []

    ppt_output_dir.mkdir(parents=True, exist_ok=True)
    outputs: list[Path] = []

    _emit(progress, "开始生成筑基臻选100练模块一 PPTX。")
    outputs.extend(
        zhuji_ppt.generate_zuhe_jiafen_pptx(
            md_run=md_run,
            template=module1_template,
            output_dir=ppt_output_dir,
            limit_groups=None,
            group_indexes=[1, 2],
            group_ranges=module1_ranges,
            progress=progress,
        )
    )

    _emit(progress, "开始生成筑基臻选100练模块二 PPTX。")
    module2_group_indexes = list(range(3, len(zhuji_ppt.GROUP_SPECS) + 1))
    outputs.extend(
        zhuji_ppt.generate_zuhe_jiafen_pptx(
            md_run=md_run,
            template=module2_template,
            output_dir=ppt_output_dir,
            limit_groups=None,
            group_indexes=module2_group_indexes,
            group_ranges=module2_ranges,
            progress=progress,
        )
    )

    manifest_path = md_output_dir / "zhuji_100_generation_manifest.json"
    _write_json(
        manifest_path,
        {
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "body_fbd": str(body_fbd),
            "answer_fbd": str(answer_fbd),
            "module1_template": str(module1_template),
            "module2_template": str(module2_template),
            "md_run": str(md_run),
            "ppt_output_dir": str(ppt_output_dir),
            "ppt_outputs": [str(path) for path in outputs],
        },
    )

    return Zhuji100GenerationResult(
        md_run=md_run,
        ppt_output_dir=ppt_output_dir,
        ppt_outputs=outputs,
        manifest_path=manifest_path,
        module1_template=module1_template,
        module2_template=module2_template,
    )

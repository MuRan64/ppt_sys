#!/usr/bin/env python3
"""Adapter for 语文《文言文阅读加分练》.

The detailed layout and matching rules are maintained in
``tests/test_wenyan_yuedu_jiafen_md_to_ppt.py``.  This adapter keeps the GUI
integration small: convert the selected FBD pair into the expected MD run and
then call the tested PPT generator for every configured group.
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

import test_wenyan_yuedu_jiafen_md_to_ppt as yuedu_ppt  # noqa: E402
from book_grouping import UserCancelled  # noqa: E402


Progress = Callable[[str], None]
GroupSelector = Callable[[list[dict[str, object]]], list[dict[str, object]] | None]


@dataclass(frozen=True)
class WenyanYueduJiafenGenerationResult:
    """Return value for GUI callers."""

    md_run: Path
    ppt_output_dir: Path
    ppt_outputs: list[Path]
    manifest_path: Path


def _emit(progress: Progress | None, message: str) -> None:
    if progress:
        progress(message)


def _write_json(path: Path, payload: dict[str, object]) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8", newline="\n")


def generate_wenyan_yuedu_jiafen_pptx_from_fbd(
    *,
    body_fbd: Path,
    answer_fbd: Path,
    template: Path,
    md_output_dir: Path,
    ppt_output_dir: Path,
    group_selector: GroupSelector | None = None,
    progress: Progress | None = None,
) -> WenyanYueduJiafenGenerationResult:
    """Full production entry point for GUI: FBD pair -> MD run -> PPTX."""

    md_run = yuedu_ppt.convert_wenyan_yuedu_fbd_pair(
        body_fbd=body_fbd,
        answer_fbd=answer_fbd,
        output_dir=md_output_dir,
        progress=progress,
    )

    group_ranges = None
    if group_selector is not None:
        units = yuedu_ppt.scan_yuedu_group_units(md_run)
        group_ranges = group_selector(units)
        if group_ranges is None:
            raise UserCancelled("用户取消了文言文阅读加分练 PPT 分组设置。")

    _emit(progress, "开始生成文言文阅读加分练 PPTX。")
    ppt_outputs = yuedu_ppt.generate_wenyan_yuedu_jiafen_pptx(
        md_run=md_run,
        template=template,
        output_dir=ppt_output_dir,
        limit_groups=None,
        group_ranges=group_ranges,
        progress=progress,
    )

    manifest_path = md_output_dir / "wenyan_yuedu_jiafen_generation_manifest.json"
    _write_json(
        manifest_path,
        {
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "body_fbd": str(body_fbd),
            "answer_fbd": str(answer_fbd),
            "template": str(template),
            "md_run": str(md_run),
            "ppt_output_dir": str(ppt_output_dir),
            "group_ranges": group_ranges,
            "ppt_outputs": [str(path) for path in ppt_outputs],
        },
    )

    return WenyanYueduJiafenGenerationResult(
        md_run=md_run,
        ppt_output_dir=ppt_output_dir,
        ppt_outputs=ppt_outputs,
        manifest_path=manifest_path,
    )

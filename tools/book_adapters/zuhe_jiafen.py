#!/usr/bin/env python3
"""Adapter for 语文《高考组合加分练》.

The tested layout rules live in ``tests/test_zuhe_jiafen_md_to_ppt.py``.  This
production-facing adapter owns the GUI orchestration: convert the source and
answer FBD files into the MD run shape expected by the PPT generator, then call
that generator for all groups.
"""

from __future__ import annotations

import json
import sys
from dataclasses import asdict, dataclass
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

from fbd_to_md import convert_fbd_to_markdown, read_fbd  # noqa: E402
from book_grouping import UserCancelled  # noqa: E402
import test_zuhe_jiafen_md_to_ppt as zuhe_ppt  # noqa: E402


Progress = Callable[[str], None]
GroupSelector = Callable[[list[dict[str, object]]], list[dict[str, object]] | None]


@dataclass(frozen=True)
class ZuheJiafenGenerationResult:
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


def _dedupe_paths(paths: list[Path]) -> list[Path]:
    result: list[Path] = []
    seen: set[str] = set()
    for path in paths:
        key = str(path.resolve() if path.exists() else path).lower()
        if key not in seen:
            result.append(path)
            seen.add(key)
    return result


def _collect_tif_keep_map(source_dirs: list[Path]) -> dict[str, str]:
    keep_map: dict[str, str] = {}
    for source_dir in _dedupe_paths(source_dirs):
        keep_map.update(zuhe_ppt.collect_tif_keep_map(source_dir))
    return keep_map


def convert_zuhe_jiafen_fbd_pair(
    *,
    body_fbd: Path,
    answer_fbd: Path,
    output_dir: Path,
    progress: Progress | None = None,
) -> Path:
    """Convert source and answer FBD files into one MD run directory."""

    output_dir.mkdir(parents=True, exist_ok=True)
    tif_source_dirs = _dedupe_paths([body_fbd.parent, answer_fbd.parent])
    tif_keep_map = _collect_tif_keep_map(tif_source_dirs)
    _emit(progress, f"高考组合加分练：自动保留 TIF 资源 {len(tif_keep_map)} 个。")

    outputs: list[dict[str, object]] = []
    for role, fbd_path, safe_name in [("正文", body_fbd, "body"), ("答案", answer_fbd, "answer")]:
        _emit(progress, f"转换{role} FBD：{fbd_path}")
        raw = read_fbd(fbd_path, "gb18030")
        markdown, stats = convert_fbd_to_markdown(raw, tif_replacements=tif_keep_map)
        md_path = output_dir / f"{role}_{safe_name}.md"
        md_path.write_text(markdown, encoding="utf-8", newline="\n")
        outputs.append({"role": role, "source": str(fbd_path), "markdown": str(md_path), "stats": asdict(stats)})
        _emit(progress, f"{role} MD：{md_path.name}，TIF 保留 {stats.tif_kept}，未知 {stats.tif_unknown}")

    manifest_path = output_dir / "conversion_manifest.json"
    _write_json(
        manifest_path,
        {
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "output_dir": str(output_dir),
            "tif_source_dirs": [str(path) for path in tif_source_dirs],
            "outputs": outputs,
            "note": "高考组合加分练：正文 FBD + 答案 FBD。",
        },
    )
    return output_dir


def generate_zuhe_jiafen_pptx_from_fbd(
    *,
    body_fbd: Path,
    answer_fbd: Path,
    template: Path,
    md_output_dir: Path,
    ppt_output_dir: Path,
    group_selector: GroupSelector | None = None,
    progress: Progress | None = None,
) -> ZuheJiafenGenerationResult:
    """Full production entry point for GUI: FBD pair -> MD run -> PPTX."""

    md_run = convert_zuhe_jiafen_fbd_pair(
        body_fbd=body_fbd,
        answer_fbd=answer_fbd,
        output_dir=md_output_dir,
        progress=progress,
    )

    group_ranges = None
    if group_selector is not None:
        units = zuhe_ppt.scan_zuhe_group_units(md_run)
        group_ranges = group_selector(units)
        if group_ranges is None:
            raise UserCancelled("用户取消了组合加分练 PPT 分组设置。")

    _emit(progress, "开始生成高考组合加分练 PPTX。")
    ppt_outputs = zuhe_ppt.generate_zuhe_jiafen_pptx(
        md_run=md_run,
        template=template,
        output_dir=ppt_output_dir,
        limit_groups=None,
        group_ranges=group_ranges,
        progress=progress,
    )

    manifest_path = md_output_dir / "zuhe_jiafen_generation_manifest.json"
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

    return ZuheJiafenGenerationResult(
        md_run=md_run,
        ppt_output_dir=ppt_output_dir,
        ppt_outputs=ppt_outputs,
        manifest_path=manifest_path,
    )

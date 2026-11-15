#!/usr/bin/env python3
"""Adapter for 语文《必背古诗文》.

This book is intentionally handled as a bundle:
- 主书题目 comes from `必背古诗文.fbd`.
- 主书答案 comes from `必背古诗文答案.fbd`.
- 小册子题目 comes from `必背古诗文册子.fbd`.
- 默写挖空练答案 is inferred from `必背口袋本.fbd`, while other booklet
  answers reuse the main answer FBD.

The heavy PPT layout rules still live in the mature test generator for this
book.  This adapter owns the production-facing orchestration so the GUI does
not need book-specific branching beyond choosing this entry point.
"""

from __future__ import annotations

import json
import re
import sys
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable


TOOLS_DIR = Path(__file__).resolve().parents[1]
PROJECT_DIR = Path(getattr(sys, "_MEIPASS", TOOLS_DIR.parent))
TESTS_DIR = PROJECT_DIR / "tests"
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))
if str(TESTS_DIR) not in sys.path:
    sys.path.insert(0, str(TESTS_DIR))

from fbd_to_md import convert_fbd_to_markdown, normalize_tif_key, read_fbd, tif_markdown_tag  # noqa: E402
from book_grouping import UserCancelled  # noqa: E402


Progress = Callable[[str], None]


@dataclass(frozen=True)
class BibeiMarkdownBundle:
    """The three MD run directories required by the PPT generation stage."""

    root_dir: Path
    main_md_run: Path
    booklet_md_run: Path
    pocket_md_run: Path


@dataclass(frozen=True)
class BibeiGenerationResult:
    """Return value for the GUI and for future command-line wrappers."""

    md_bundle: BibeiMarkdownBundle
    ppt_output_dir: Path
    ppt_outputs: list[Path]
    manifest_path: Path


def _emit(progress: Progress | None, message: str) -> None:
    if progress:
        progress(message)


def _initial_ranges(groups: list[list[dict[str, Any]]]) -> list[dict[str, Any]]:
    ranges: list[dict[str, Any]] = []
    for members in groups:
        if not members:
            continue
        volumes = {str(unit["volume"]) for unit in members}
        ranges.append(
            {
                "volume": next(iter(volumes)) if len(volumes) == 1 else "跨章节",
                "unit_keys": [(str(unit["volume"]), int(unit["num"])) for unit in members],
                "start": int(members[0]["num"]),
                "end": int(members[-1]["num"]),
            }
        )
    return ranges


def _dedupe_paths(paths: list[Path]) -> list[Path]:
    result: list[Path] = []
    seen: set[str] = set()
    for path in paths:
        key = str(path.resolve() if path.exists() else path).lower()
        if key not in seen:
            result.append(path)
            seen.add(key)
    return result


def collect_tif_keep_map(source_dirs: list[Path]) -> dict[str, str]:
    """Keep every TIF referenced by this book.

    Earlier Chinese books asked the user to confirm TIFs one by one.  For this
    book, the source folders have already been cleaned so the remaining TIFs are
    meaningful glyphs, exercise images, or title assets.
    """

    keep_map: dict[str, str] = {}
    for source_dir in _dedupe_paths(source_dirs):
        if not source_dir.exists():
            continue
        for path in sorted(source_dir.rglob("*.tif")) + sorted(source_dir.rglob("*.tiff")):
            if path.is_file():
                keep_map[normalize_tif_key(path.name)] = tif_markdown_tag(path.name)
    return keep_map


def preprocess_pocket_raw_text(raw: str) -> tuple[str, dict[str, int]]:
    """Remove pocket-book spacing placeholders before semantic conversion."""

    placeholder_count = raw.count("　")
    return raw.replace("　", ""), {"spacing_placeholders_removed": placeholder_count}


def tidy_pocket_markdown(markdown: str) -> str:
    """Collapse visual spacing that the pocket book uses inside Chinese text."""

    return re.sub(r"(?<=[\u4e00-\u9fff]) (?=[\u4e00-\u9fff])", "", markdown)


def _write_json(path: Path, payload: dict[str, object]) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8", newline="\n")


def _convert_one_fbd(
    *,
    role: str,
    fbd_path: Path,
    run_dir: Path,
    tif_replacements: dict[str, str],
    start_marker: str | None = None,
    preprocess_pocket: bool = False,
) -> tuple[Path, dict[str, object]]:
    raw = read_fbd(fbd_path, "gb18030")
    preprocess_stats: dict[str, int] = {}
    if preprocess_pocket:
        raw, preprocess_stats = preprocess_pocket_raw_text(raw)

    markdown, stats = convert_fbd_to_markdown(
        raw,
        start_marker=start_marker,
        tif_replacements=tif_replacements,
    )
    if preprocess_pocket:
        markdown = tidy_pocket_markdown(markdown)

    md_path = run_dir / f"{role}_{fbd_path.stem}_格式保留.md"
    md_path.write_text(markdown, encoding="utf-8", newline="\n")
    return md_path, {"stats": asdict(stats), "preprocess": preprocess_stats}


def _write_run_manifest(
    *,
    run_dir: Path,
    tif_source_dirs: list[Path],
    tif_keep_map: dict[str, str],
    outputs: list[dict[str, object]],
    note: str,
) -> Path:
    tif_map_path = run_dir / "必背古诗文_tif_map.json"
    _write_json(tif_map_path, {"keep": sorted(tif_keep_map), "ignore": [], "source_dirs": [str(path) for path in tif_source_dirs]})

    manifest = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "output_dir": str(run_dir),
        "tif_map": str(tif_map_path),
        # PPT image lookup reads both `outputs[*].source` and `tif_source_dirs`.
        "tif_source_dirs": [str(path) for path in tif_source_dirs],
        "outputs": outputs,
        "note": note,
    }
    manifest_path = run_dir / "conversion_manifest.json"
    _write_json(manifest_path, manifest)
    return manifest_path


def convert_bibei_fbd_bundle(
    *,
    body_fbd: Path,
    answer_fbd: Path,
    booklet_fbd: Path,
    pocket_fbd: Path,
    output_dir: Path,
    progress: Progress | None = None,
) -> BibeiMarkdownBundle:
    """Convert the four 必背古诗文 FBD files into the MD layout expected by PPT.

    The booklet run deliberately contains the booklet body plus the main answer
    MD, because the small-booklet “重点语句强化练” answers live in the main answer
    source.  The pocket run stays separate and is passed into the generator to
    infer answers for “默写挖空练”.
    """

    output_dir.mkdir(parents=True, exist_ok=True)
    main_run = output_dir / "主书_md"
    booklet_run = output_dir / "小册子_md"
    pocket_run = output_dir / "口袋本_md"
    for run_dir in (main_run, booklet_run, pocket_run):
        run_dir.mkdir(parents=True, exist_ok=True)

    tif_source_dirs = _dedupe_paths([body_fbd.parent, answer_fbd.parent, booklet_fbd.parent, pocket_fbd.parent])
    tif_keep_map = collect_tif_keep_map(tif_source_dirs)
    _emit(progress, f"必背古诗文：自动保留 TIF 资源 {len(tif_keep_map)} 个。")

    main_outputs: list[dict[str, object]] = []
    for role, fbd_path in [("正文", body_fbd), ("答案", answer_fbd)]:
        _emit(progress, f"转换主书{role}：{fbd_path}")
        md_path, detail = _convert_one_fbd(role=role, fbd_path=fbd_path, run_dir=main_run, tif_replacements=tif_keep_map)
        main_outputs.append({"role": role, "source": str(fbd_path), "markdown": str(md_path), **detail})
    _write_run_manifest(
        run_dir=main_run,
        tif_source_dirs=tif_source_dirs,
        tif_keep_map=tif_keep_map,
        outputs=main_outputs,
        note="必背古诗文主书：正文 FBD + 答案 FBD。",
    )

    booklet_outputs: list[dict[str, object]] = []
    for role, fbd_path in [("正文", booklet_fbd), ("答案", answer_fbd)]:
        _emit(progress, f"转换小册子{role}：{fbd_path}")
        md_path, detail = _convert_one_fbd(role=role, fbd_path=fbd_path, run_dir=booklet_run, tif_replacements=tif_keep_map)
        booklet_outputs.append({"role": role, "source": str(fbd_path), "markdown": str(md_path), **detail})
    _write_run_manifest(
        run_dir=booklet_run,
        tif_source_dirs=tif_source_dirs,
        tif_keep_map=tif_keep_map,
        outputs=booklet_outputs,
        note="必背古诗文小册子：册子正文 FBD + 主书答案 FBD。",
    )

    _emit(progress, f"转换口袋本：{pocket_fbd}")
    pocket_md, pocket_detail = _convert_one_fbd(
        role="口袋本",
        fbd_path=pocket_fbd,
        run_dir=pocket_run,
        tif_replacements=tif_keep_map,
        start_marker="短歌行",
        preprocess_pocket=True,
    )
    _write_run_manifest(
        run_dir=pocket_run,
        tif_source_dirs=tif_source_dirs,
        tif_keep_map=tif_keep_map,
        outputs=[{"role": "口袋本", "source": str(pocket_fbd), "markdown": str(pocket_md), **pocket_detail}],
        note="必背口袋本：提供完整诗文，用于小册子默写挖空练反推答案。",
    )

    return BibeiMarkdownBundle(
        root_dir=output_dir,
        main_md_run=main_run,
        booklet_md_run=booklet_run,
        pocket_md_run=pocket_run,
    )


def generate_bibei_book_pptx_from_fbd(
    *,
    body_fbd: Path,
    answer_fbd: Path,
    booklet_fbd: Path,
    pocket_fbd: Path,
    template: Path,
    md_output_dir: Path,
    ppt_output_dir: Path,
    group_selector: Callable[
        [str, list[dict[str, Any]], list[dict[str, Any]]], list[dict[str, Any]] | None
    ] | None = None,
    progress: Progress | None = None,
) -> BibeiGenerationResult:
    """Full production entry point for GUI: FBD bundle -> MD bundle -> PPTX."""

    bundle = convert_bibei_fbd_bundle(
        body_fbd=body_fbd,
        answer_fbd=answer_fbd,
        booklet_fbd=booklet_fbd,
        pocket_fbd=pocket_fbd,
        output_dir=md_output_dir,
        progress=progress,
    )

    # Lazy import avoids changing the shared Wenyan PPT hooks unless this
    # adapter is actually selected.
    import test_bibei_gushiwen_md_to_ppt as bibei_ppt  # noqa: E402

    main_group_ranges = None
    booklet_group_ranges = None
    if group_selector is not None:
        main_units = bibei_ppt.scan_bibei_main_group_units(bundle.main_md_run)
        _emit(progress, f"主书正文 MD 扫描到 {len(main_units)} 个篇目/练习，等待设置主书目录分组。")
        core_units = [
            unit
            for unit in main_units
            if 1 <= int(unit["num"]) <= 60
            and unit["volume"] != "附赠选修12篇"
            and unit["title"] != "150组冲刺练"
        ]
        extra_units = [unit for unit in main_units if unit["volume"] == "附赠选修12篇"]
        rush_units = [unit for unit in main_units if unit["title"] == "150组冲刺练"]
        core_groups = [core_units[index : index + 6] for index in range(0, len(core_units), 6)]
        main_defaults = _initial_ranges([*core_groups, extra_units, *[[unit] for unit in rush_units]])
        main_group_ranges = group_selector("主书", main_units, main_defaults)
        if main_group_ranges is None:
            raise UserCancelled("用户取消了《必背古诗文》主书 PPT 分组设置。")

        booklet_units = bibei_ppt.scan_bibei_booklet_group_units(
            bundle.booklet_md_run,
            booklet_open_source_dir=booklet_fbd.parent,
            booklet_open_image_md_run=bundle.root_dir / "小册子开放性图片_md",
        )
        _emit(progress, f"小册子正文 MD 扫描到 {len(booklet_units)} 个分册/专题，等待设置小册子目录分组。")
        booklet_sections: dict[str, list[dict[str, Any]]] = {}
        for unit in booklet_units:
            booklet_sections.setdefault(str(unit["volume"]), []).append(unit)
        booklet_defaults = _initial_ranges(
            [members if volume == "开放性默写专题练" else members[:1] for volume, members in booklet_sections.items()]
        )
        booklet_group_ranges = group_selector("小册子", booklet_units, booklet_defaults)
        if booklet_group_ranges is None:
            raise UserCancelled("用户取消了《必背古诗文》小册子 PPT 分组设置。")

    ppt_output_dir.mkdir(parents=True, exist_ok=True)
    outputs: list[Path] = []

    _emit(progress, "开始生成必背古诗文主书 PPTX。")
    outputs.extend(
        bibei_ppt.generate_bibei_main_complete_pptx(
            main_md_run=bundle.main_md_run,
            template=template,
            output_root=ppt_output_dir,
            group_ranges=main_group_ranges,
            progress=progress,
        )
    )

    _emit(progress, "开始生成必背古诗文小册子 PPTX。")
    outputs.extend(
        bibei_ppt.generate_bibei_booklet_pptx(
            booklet_md_run=bundle.booklet_md_run,
            pocket_md_run=bundle.pocket_md_run,
            template=template,
            output_dir=ppt_output_dir / "小册子",
            group_ranges=booklet_group_ranges,
            booklet_open_source_dir=booklet_fbd.parent,
            booklet_open_image_md_run=bundle.root_dir / "小册子开放性图片_md",
            progress=progress,
        )
    )

    manifest_path = md_output_dir / "bibei_generation_manifest.json"
    _write_json(
        manifest_path,
        {
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "body_fbd": str(body_fbd),
            "answer_fbd": str(answer_fbd),
            "booklet_fbd": str(booklet_fbd),
            "pocket_fbd": str(pocket_fbd),
            "template": str(template),
            "md_root": str(bundle.root_dir),
            "ppt_output_dir": str(ppt_output_dir),
            "main_group_ranges": main_group_ranges,
            "booklet_group_ranges": booklet_group_ranges,
            "ppt_outputs": [str(path) for path in outputs],
        },
    )

    return BibeiGenerationResult(
        md_bundle=bundle,
        ppt_output_dir=ppt_output_dir,
        ppt_outputs=outputs,
        manifest_path=manifest_path,
    )


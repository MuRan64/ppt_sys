#!/usr/bin/env python3
"""Test generator for 语文《筑基臻选100练》.

This adapter starts from the mature 《组合加分练》 PPT pipeline because this book
shares most question forms with it.  Book-specific logic is kept here: FBD/MD
paths, module/practice splitting, and the output grouping for 筑基臻选100练.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import shutil
import subprocess
import sys
import zipfile
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Callable


PROJECT_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
TESTS_DIR = Path(__file__).resolve().parent
TOOLS_DIR = PROJECT_DIR / "tools"
if str(TESTS_DIR) not in sys.path:
    sys.path.insert(0, str(TESTS_DIR))
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

import test_md_to_ppt as wenyan  # noqa: E402
from fbd_to_md import convert_fbd_to_markdown, normalize_tif_key, read_fbd, tif_markdown_tag  # noqa: E402
from generation_state import isolated_generation  # noqa: E402


DEFAULT_SOURCE_DIR = PROJECT_DIR / "筑基臻选100练_源文件"
DEFAULT_BODY_FBD = DEFAULT_SOURCE_DIR / "筑基100练 正文.fbd"
DEFAULT_ANSWER_FBD = DEFAULT_SOURCE_DIR / "筑基100练 答案.fbd"
DEFAULT_TEMPLATE_MODULE1 = PROJECT_DIR / "筑基模块一模板.pptx"
DEFAULT_TEMPLATE_MODULE2 = PROJECT_DIR / "筑基模块二模板.pptx"
DEFAULT_TEMPLATE = DEFAULT_TEMPLATE_MODULE1
DEFAULT_MD_ROOT = PROJECT_DIR / "out_md" / "筑基臻选100练"
DEFAULT_OUTPUT_DIR = PROJECT_DIR / "output" / "test_zhuji_100_ppt"

GROUP_SPECS = [
    (r"模块一 语用专项40练\1字词+语病+标点+修辞+表达效果（1~20练）.pptx", "语用专项40练", 1, 20),
    (r"模块一 语用专项40练\2衔接+图表+压缩+语言表达（21~40练）.pptx", "语用专项40练", 21, 40),
    (r"模块二 组合通关60练\①一年新题通关\1模块二 一年新题通关41~45练.pptx", "一年新题通关", 41, 45),
    (r"模块二 组合通关60练\①一年新题通关\2模块二 一年新题通关46~50练.pptx", "一年新题通关", 46, 50),
    (r"模块二 组合通关60练\①一年新题通关\3模块二 一年新题通关51~55练.pptx", "一年新题通关", 51, 55),
    (r"模块二 组合通关60练\①一年新题通关\4模块二 一年新题通关56~60练.pptx", "一年新题通关", 56, 60),
    (r"模块二 组合通关60练\②三年模拟题通关\1模块二 三年模拟题通关61~65练.pptx", "三年模拟题通关", 61, 65),
    (r"模块二 组合通关60练\②三年模拟题通关\2模块二 三年模拟题通关66~70练.pptx", "三年模拟题通关", 66, 70),
    (r"模块二 组合通关60练\②三年模拟题通关\3模块二 三年模拟题通关71~75练.pptx", "三年模拟题通关", 71, 75),
    (r"模块二 组合通关60练\②三年模拟题通关\4模块二 三年模拟题通关76~80练.pptx", "三年模拟题通关", 76, 80),
    (r"模块二 组合通关60练\③创新预测题通关\1模块二 创新预测题题通关81~85练.pptx", "创新预测题通关", 81, 85),
    (r"模块二 组合通关60练\③创新预测题通关\2模块二 创新预测题通关86~90练.pptx", "创新预测题通关", 86, 90),
    (r"模块二 组合通关60练\③创新预测题通关\3模块二 创新预测题通关91~95练.pptx", "创新预测题通关", 91, 95),
    (r"模块二 组合通关60练\③创新预测题通关\4模块二 创新预测题通关96~100练.pptx", "创新预测题通关", 96, 100),
]
ZHUJI_MODULE2_GROUPS = {"一年新题通关", "三年模拟题通关", "创新预测题通关", "组合通关60练"}

SECTION_RE = re.compile(r"^[一二三四五六七八九十]+(?:、|[\s　]+)")
PRACTICE_TITLE_RE = re.compile(r"^第?(\d{1,3})练(?:[:：].*)?$")
ZUHE_TEAL = "3A7371"
ZUHE_BLUE = "00AEEF"
ZUHE_SUMMARY_ORDER = (
    "字词强化练",
    "语病强化练",
    "标点强化练",
    "句式强化练",
    "修辞强化练",
    "表达效果强化练",
)
ZUHE_ANALYSIS_MARKERS = ("【关键能力】", "【解题思路】", "【解决思路】", "【解析】")
ZUHE_REFERENCE_TIF_NAMES = (
    "参考译文.tif",
    "诗歌鉴赏.tif",
    "诗词鉴赏.tif",
    "知识拓展.tif",
)
ZUHE_LONG_ANSWER_TEXT_UNITS = 30.0
ZUHE_LONG_ANSWER_LINE_UNITS = 54.0
ZUHE_LONG_ANSWER_UNDERLINE_UNITS = 34.0
ZUHE_LONG_ANSWER_TRAILING_UNITS = 3.0
ZUHE_REVISION_INLINE_UNITS = 20.0
ZUHE_REVISION_LINE_UNITS = 28.0
ZUHE_ANALYSIS_FRAGMENT_UNITS = 300.0
ZUHE_ANALYSIS_SINGLE_BLOCK_MAX_LINES = 13
ZHUJI_ANSWER_VERTICAL_SHIFT_PT = 4.0
ZUHE_MAJOR_SECTION_KEYWORDS = (
    "字词强化练",
    "语病强化练",
    "标点强化练",
    "句式强化练",
    "修辞强化练",
    "表达效果强化练",
    "语言文字运用综合练",
    "现代文阅读强化练",
    "古代诗文阅读强化练",
    "写作强化练",
    "一年新题通关",
    "三年模拟题通关",
    "创新预测题通关",
)
ZUHE_MATERIAL_LABEL_RE = re.compile(r"^((?:材料|文本)[一二三四五六七八九十]+[：:])\s*(.*)$")

ZHUJI_MODULE_TITLES = (
    "字词强化练",
    "语病强化练",
    "标点强化练",
    "句式强化练",
    "修辞强化练",
    "表达效果强化练",
    "语言文字运用综合练",
    "现代文阅读强化练",
    "古代诗文阅读强化练",
    "写作强化练",
    "一年新题通关",
    "三年模拟题通关",
    "创新预测题通关",
)
ZHUJI_BOOK_TITLES = {
    "高考语文筑基臻选100练",
    "语用专项40练",
    "组合通关60练",
    "现代文阅读专项30练",
    "古代诗文阅读专项20练",
    "写作专项10练",
}
ZHUJI_TOP_DIRECTORY_ITEMS = (
    "字词强化练",
    "语病强化练",
    "标点、句式、修辞强化练",
    "表达效果强化练",
)
ZHUJI_TOP_DIRECTORY_GROUPS = (
    ("字词强化练", ("字词强化练",)),
    ("语病强化练", ("语病强化练",)),
    ("标点、句式、修辞强化练", ("标点强化练", "句式强化练", "修辞强化练")),
    ("表达效果强化练", ("表达效果强化练",)),
)
ZHUJI_INLINE_SECTION_TITLES = {
    "语言文字运用",
    "文言一词多义",
    "名篇名句默写",
}


def plain(text: str) -> str:
    return wenyan.strip_inline_tags(text).strip()


def marker(text: str) -> str:
    return re.sub(r"[\s　]+", "", plain(text))


def has_fbd_tif(line: str) -> bool:
    return "data-fbd-tif" in line or bool(re.search(r"\.(?:tif|tiff)\b", line, flags=re.I))


def is_zuhe_reference_tif_line(line: str) -> bool:
    lower = line.lower()
    return has_fbd_tif(line) and any(name.lower() in lower for name in ZUHE_REFERENCE_TIF_NAMES)


def mark_zuhe_wave_blue(line: str) -> str:
    """Keep source reading underlines blue after later line splitting.

    Answer blanks are generated later as ``answer-fill uline`` and must stay
    black, so this only touches raw source ``uline`` / ``wave`` spans.
    """

    def repl(match: re.Match[str]) -> str:
        classes = match.group(1)
        class_set = classes.split()
        if "answer-fill" in class_set or "blue-underline" in class_set:
            return match.group(0)
        if "wave" not in class_set and "uline" not in class_set:
            return match.group(0)
        class_set.append("blue-underline")
        return f'<span class="{" ".join(class_set)}"'

    return re.sub(r'<span\s+class="([^"]*\b(?:wave|uline)\b[^"]*)"', repl, line)


def zuhe_material_label_parts(line: str) -> tuple[str, str] | None:
    match = ZUHE_MATERIAL_LABEL_RE.match(plain(line))
    if not match:
        return None
    return match.group(1), match.group(2).strip()


def style_zuhe_material_label(line: str) -> tuple[str, str]:
    parts = zuhe_material_label_parts(line)
    if not parts:
        return line, "extended_reading"
    label, rest = parts
    if not rest:
        return line, "subheading"
    escaped_label = re.escape(label)
    styled = re.sub(
        escaped_label,
        f'<span class="font-heiti bold-text">{label}</span>',
        line,
        count=1,
    )
    return styled, "material_reading"


def style_zuhe_analysis_markers(text: str) -> str:
    for marker in ZUHE_ANALYSIS_MARKERS:
        text = text.replace(marker, f'<span class="font-heiti bold-text">{marker}</span>')
    return text


def is_zuhe_source_note_plain(text: str) -> bool:
    if wenyan.is_source_note_plain(text):
        return True
    return bool(re.match(r"^[（(].*(节选自|摘编自|摘译自|选自|有删改)", text))


def is_zuhe_major_section_title(line: str) -> bool:
    text = plain(line)
    if wenyan.is_section_marker(text):
        return True
    return bool(
        SECTION_RE.match(text)
        and (
            any(keyword in text for keyword in ZUHE_MAJOR_SECTION_KEYWORDS)
            or any(keyword in text for keyword in ZHUJI_INLINE_SECTION_TITLES)
        )
    )


def is_zuhe_minor_section_title(line: str) -> bool:
    text = plain(line)
    compact = marker(line)
    return bool(
        re.match(r"^[（(][一二三四五六七八九十]+[）)]", compact)
        and any(keyword in text for keyword in ZUHE_MAJOR_SECTION_KEYWORDS)
    )


def zuhe_split_sections_with_context(piece: wenyan.Piece) -> list[tuple[str, str, list[str]]]:
    sections: list[tuple[str, str, list[str]]] = []
    current_title = "导入"
    current_lines: list[str] = []

    source = piece.lines[1:]
    if piece.author and source and source[0] == piece.author:
        source = source[1:]

    for line in source:
        if line in wenyan.MODULE_TITLES or is_zuhe_major_section_title(line):
            if current_lines:
                sections.append(("", current_title, current_lines))
            current_title = line
            current_lines = [line]
        else:
            current_lines.append(line)
    if current_lines:
        sections.append(("", current_title, current_lines))
    return sections


def collect_tif_keep_map(source_dir: Path) -> dict[str, str]:
    keep_map: dict[str, str] = {}
    for path in sorted(source_dir.rglob("*.tif")) + sorted(source_dir.rglob("*.tiff")):
        if path.is_file():
            keep_map[normalize_tif_key(path.name)] = tif_markdown_tag(path.name)
    return keep_map


def convert_zuhe_fbd_to_md(
    *,
    body_fbd: Path = DEFAULT_BODY_FBD,
    answer_fbd: Path = DEFAULT_ANSWER_FBD,
    output_root: Path = DEFAULT_MD_ROOT,
    progress: Callable[[str], None] | None = None,
) -> Path:
    """Convert source/answer FBD files into one Markdown run directory."""

    def emit(message: str) -> None:
        if progress:
            progress(message)
        else:
            print(message)

    source_dir = body_fbd.parent
    tif_keep_map = collect_tif_keep_map(source_dir)
    run_dir = output_root / datetime.now().strftime("run_%Y%m%d_%H%M%S")
    run_dir.mkdir(parents=True, exist_ok=False)

    outputs: list[dict[str, object]] = []
    for role, fbd_path, safe_name in [("正文", body_fbd, "body"), ("答案", answer_fbd, "answer")]:
        emit(f"转换{role} FBD：{fbd_path}")
        raw = read_fbd(fbd_path, "gb18030")
        markdown, stats = convert_fbd_to_markdown(raw, tif_replacements=tif_keep_map)
        md_path = run_dir / f"{role}_{safe_name}.md"
        md_path.write_text(markdown, encoding="utf-8", newline="\n")
        outputs.append({"role": role, "source": str(fbd_path), "markdown": str(md_path), "stats": asdict(stats)})
        emit(f"{role} MD：{md_path.name}，TIF 保留 {stats.tif_kept}，未知 {stats.tif_unknown}")

    manifest = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "output_dir": str(run_dir),
        "tif_source_dirs": [str(source_dir)],
        "outputs": outputs,
    }
    (run_dir / "conversion_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
        newline="\n",
    )
    return run_dir


def find_md_files(run_dir: Path) -> tuple[Path, Path]:
    body = sorted(run_dir.glob("正文_*.md"))
    answer = sorted(run_dir.glob("答案_*.md"))
    if not body or not answer:
        raise FileNotFoundError(f"未在目录中找到正文_*.md 和 答案_*.md：{run_dir}")
    return body[0], answer[0]


def is_group_title(text: str) -> bool:
    return plain(text) in ZHUJI_MODULE_TITLES


def zhuji_practice_title_match(line: str) -> re.Match[str] | None:
    return PRACTICE_TITLE_RE.fullmatch(marker(line))


def is_practice_start(lines: list[str], index: int) -> re.Match[str] | None:
    return zhuji_practice_title_match(lines[index])


def normalize_section_scores(lines: list[str]) -> list[str]:
    result: list[str] = []
    index = 0
    while index < len(lines):
        line = lines[index]
        text = plain(line)
        if is_zuhe_major_section_title(line) and index + 1 < len(lines) and plain(lines[index + 1]).isdigit():
            result.append(f"{text}「{plain(lines[index + 1])}分」")
            index += 2
            continue
        result.append(line)
        index += 1
    return result


def replace_choice_layout_blanks(line: str) -> str:
    """Choice-option blanks are visual spacing from FBD, not answer blanks."""

    probe = re.sub(r"^[\s_＿]+", "", plain(line))
    if not re.match(r"^[A-D][\.．、]", probe):
        return line

    def blank_repl(match: re.Match[str]) -> str:
        width_match = re.search(r"width\s*:\s*([0-9.]+)em", match.group(0))
        count = int(round(float(width_match.group(1)))) if width_match else 2
        return "　" * max(1, count)

    def underscore_repl(match: re.Match[str]) -> str:
        # clean_lines() has already converted blank spans to underscores by the
        # time this function runs in normal parsing.
        count = int(round(len(match.group(0)) / 1.8))
        return "　" * max(1, count)

    line = re.sub(r'<span\s+class="blank"[^>]*></span>', blank_repl, line)
    return re.sub(r"[_＿]{3,}", underscore_repl, line)


def remove_body_noise(lines: list[str], group: str) -> list[str]:
    cleaned: list[str] = []
    for line in normalize_section_scores(lines):
        text = plain(line)
        compact = marker(line)
        if not text and not has_fbd_tif(line):
            continue
        if text == group or text in ZHUJI_MODULE_TITLES:
            continue
        if text in ZHUJI_BOOK_TITLES or text in {"分"}:
            continue
        if re.fullmatch(r"\d{1,3}", text):
            continue
        if compact in {
            "用时：分钟满分：分得分：",
            "总分：分推荐用时：分钟",
            "第练用时：分钟满分：分得分：",
        }:
            continue
        cleaned.append(replace_choice_layout_blanks(line))
    return cleaned


def merge_zhuji_inline_section_titles(lines: list[str]) -> list[str]:
    """Join FBD-split module headings such as "二" + "文言一词多义"."""

    merged: list[str] = []
    index = 0
    while index < len(lines):
        current = plain(lines[index])
        if (
            re.fullmatch(r"[一二三四五六七八九十]+", current)
            and index + 1 < len(lines)
            and plain(lines[index + 1]) in ZHUJI_INLINE_SECTION_TITLES
        ):
            merged.append(f"{current} {plain(lines[index + 1])}")
            index += 2
            continue
        merged.append(lines[index])
        index += 1
    return merged


def zhuji_promote_module2_subtitle(title: str, block: list[str]) -> tuple[str, list[str]]:
    """Module two writes the practice type on the line after "第n练"."""

    if "：" in title or ":" in title or not block:
        return title, block
    first = plain(block[0])
    if "+" not in first:
        return title, block
    if wenyan.text_units(first) > 30 or re.search(r"[。？！；]", first):
        return title, block
    return f"{title}：{first}", block[1:]


def parse_body_pieces(md_path: Path) -> list[wenyan.Piece]:
    lines = [line for line in wenyan.clean_lines(md_path.read_text(encoding="utf-8")) if line.strip()]
    starts: list[tuple[int, str, int, str]] = []
    current_group = "语用专项40练"
    for index, line in enumerate(lines):
        if is_group_title(line):
            current_group = plain(line)
            continue
        match = is_practice_start(lines, index)
        if match:
            starts.append((index, current_group, int(match.group(1)), plain(line)))

    pieces: list[wenyan.Piece] = []
    for pos, (start, group, num, title) in enumerate(starts):
        end = starts[pos + 1][0] if pos + 1 < len(starts) else len(lines)
        block = mark_zuhe_body_lines(merge_zhuji_inline_section_titles(remove_body_noise(lines[start + 1 : end], group)))
        title, block = zhuji_promote_module2_subtitle(title, block)
        if not block:
            continue
        pieces.append(wenyan.Piece(volume=group, num=num, title=title, author="", lines=[title, *block]))
    return pieces


def scan_zhuji_group_units(md_run: Path) -> list[dict[str, object]]:
    body_md, _answer_md = find_md_files(md_run)
    return [
        {"volume": piece.volume, "num": piece.num, "title": piece.title}
        for piece in parse_body_pieces(body_md)
    ]


def parse_answer_piece_blocks(md_path: Path) -> dict[tuple[str, int], list[str]]:
    lines = [line for line in wenyan.clean_lines(md_path.read_text(encoding="utf-8")) if line.strip()]
    current_group = "语用专项40练"
    starts: list[tuple[int, str, int]] = []
    for index, line in enumerate(lines):
        if is_group_title(line):
            current_group = plain(line)
            continue
        match = zhuji_practice_title_match(line)
        if match:
            starts.append((index, current_group, int(match.group(1))))

    blocks: dict[tuple[str, int], list[str]] = {}
    for pos, (start, group, num) in enumerate(starts):
        end = starts[pos + 1][0] if pos + 1 < len(starts) else len(lines)
        block = [
            line
            for line in lines[start + 1 : end]
            if (plain(line) or has_fbd_tif(line)) and plain(line) != group and plain(line) not in ZHUJI_MODULE_TITLES
        ]
        blocks[(group, num)] = block
    return blocks


def answer_line_question_num(line: str) -> int | None:
    match = wenyan.answer_item_match(line)
    return int(match.group(1)) if match else None


def answer_blocks_by_question(lines: list[str]) -> dict[int, list[str]]:
    result: dict[int, list[str]] = {}
    current_num: int | None = None
    for line in lines:
        num = answer_line_question_num(line)
        if num is not None:
            current_num = num
            result[current_num] = [line]
            continue
        if current_num is not None:
            result[current_num].append(line)
    return result


def ordered_answer_blocks_by_question(lines: list[str]) -> list[tuple[int, list[str]]]:
    blocks: list[tuple[int, list[str]]] = []
    current_num: int | None = None
    current_lines: list[str] = []
    for line in lines:
        num = answer_line_question_num(line)
        if num is not None:
            if current_num is not None:
                blocks.append((current_num, current_lines))
            current_num = num
            current_lines = [line]
            continue
        if current_num is not None:
            current_lines.append(line)
    if current_num is not None:
        blocks.append((current_num, current_lines))
    return blocks


def renumber_answer_block(block: list[str], new_num: int) -> list[str]:
    if not block:
        return []
    result = list(block)
    result[0] = re.sub(r"^\d+([\.．、])", rf"{new_num}\1", result[0], count=1)
    return result


def is_zhuji_unnumbered_choice_section(section_lines: list[str]) -> bool:
    """Handle one-off choice practices whose stem has no leading question number."""

    plain_lines = [plain(line).strip() for line in section_lines if plain(line).strip() or has_fbd_tif(line)]
    if not plain_lines:
        return False
    has_stem_blank = any(
        re.search(r"[（(]\s*(?:[_＿]{2,}|<span\b[^>]*class=\"blank\"[^>]*>)\s*[）)]", line)
        for line in section_lines
    )
    option_letters = {
        text[0]
        for text in plain_lines
        if re.match(r"^[A-D][\.．、]", text)
    }
    return has_stem_blank and {"A", "B", "C", "D"}.issubset(option_letters)


def zhuji_answer_block_is_bare_choice(answer_block: list[str] | None) -> bool:
    if not answer_block:
        return False
    first = ""
    for line in answer_block:
        text = plain(line).strip()
        if text:
            first = text
            break
    return bool(re.fullmatch(r"[A-D]", first))


def body_section_question_nums(section_lines: list[str]) -> list[int]:
    nums: list[int] = []
    for line in section_lines:
        match = wenyan.body_item_match(line)
        if match:
            nums.append(int(match.group(1)))
    return nums


def zhuji_answer_block_payload(block: list[str] | None) -> str:
    if not block:
        return ""
    payload: list[str] = []
    for index, line in enumerate(block):
        text = line.strip()
        if index == 0:
            match = wenyan.answer_item_match(text)
            if match:
                text = match.group(2)
        text = plain(text)
        if text:
            payload.append(text)
    return "\n".join(payload)


def zhuji_answer_block_circled_values(block: list[str] | None) -> list[str]:
    if not block:
        return []
    first = block[0].strip()
    match = wenyan.answer_item_match(first)
    if match:
        first = match.group(2)
    first = plain(first)
    if not re.fullmatch(rf"[{re.escape(wenyan.CIRCLED_NUMBERS)}\s,，、；;]+", first):
        return []
    return re.findall(f"[{re.escape(wenyan.CIRCLED_NUMBERS)}]", first)


def zhuji_answer_block_looks_vocab(block: list[str] | None, row_count: int | None = None) -> bool:
    values = zhuji_answer_block_circled_values(block)
    if not values:
        return False
    return row_count is None or len(values) == row_count


def zhuji_answer_block_looks_memorization(block: list[str] | None) -> bool:
    payload = zhuji_answer_block_payload(block)
    return bool(re.search(r"[（(]1[）)]", payload)) and bool(re.search(r"[（(]2[）)]", payload))


def zhuji_vocab_table_row_count(group_lines: list[str]) -> int | None:
    table_lines, _tail_lines = zhuji_split_vocab_group_tail(group_lines)
    cleaned = [line.strip() for line in table_lines if line.strip()]
    if not is_zhuji_vocab_table_item(cleaned):
        return None
    plain_lines = [plain(line) for line in cleaned]
    try:
        example_index = plain_lines.index("例句义项")
        link_index = plain_lines.index("高考链接")
    except ValueError:
        return None
    examples = [line for line in cleaned[example_index + 1 : link_index] if plain(line)]
    links = [line for line in cleaned[link_index + 1 :] if plain(line)]
    return len(examples) + len(links)


def zhuji_resolve_vocab_answer_blocks(
    section_lines: list[str],
    by_num: dict[int, list[str]],
    ordered: list[tuple[int, list[str]]],
    used_answer_nums: set[int],
) -> dict[int, tuple[int, list[str]]]:
    groups: list[tuple[int, list[str], int]] = []
    for kind, num, group_lines in zuhe_split_body_numbered_segments(section_lines):
        if kind != "group" or num is None:
            continue
        row_count = zhuji_vocab_table_row_count(group_lines)
        if row_count is not None:
            groups.append((num, group_lines, row_count))
    if not groups:
        return {}

    result: dict[int, tuple[int, list[str]]] = {}
    ordered_index = {num: index for index, (num, _block) in enumerate(ordered)}
    cursor_index: int | None = None
    sequence_mode = False

    for body_num, _group_lines, row_count in groups:
        current = by_num.get(body_num)
        current_ok = body_num not in used_answer_nums and zhuji_answer_block_looks_vocab(current, row_count)
        if not sequence_mode and current_ok:
            result[body_num] = (body_num, current or [])
            used_answer_nums.add(body_num)
            cursor_index = ordered_index.get(body_num, cursor_index)
            continue

        candidates: list[tuple[int, int, list[str]]] = []
        for index, (answer_num, block) in enumerate(ordered):
            if answer_num in used_answer_nums:
                continue
            if sequence_mode and cursor_index is not None and index <= cursor_index:
                continue
            if zhuji_answer_block_looks_vocab(block, row_count):
                candidates.append((index, answer_num, block))

        if candidates:
            if not sequence_mode:
                candidates.sort(key=lambda item: (abs(item[1] - body_num), 0 if item[1] < body_num else 1, item[0]))
            index, answer_num, block = candidates[0]
            result[body_num] = (answer_num, renumber_answer_block(block, body_num))
            used_answer_nums.add(answer_num)
            cursor_index = index
            sequence_mode = sequence_mode or answer_num != body_num or not current_ok
            continue

        if current:
            result[body_num] = (body_num, current)
            used_answer_nums.add(body_num)

    return result


def zhuji_resolve_memorization_answer_blocks(
    section_lines: list[str],
    by_num: dict[int, list[str]],
    ordered: list[tuple[int, list[str]]],
    used_answer_nums: set[int],
) -> dict[int, tuple[int, list[str]]]:
    result: dict[int, tuple[int, list[str]]] = {}
    for body_num in body_section_question_nums(section_lines):
        current = by_num.get(body_num)
        if body_num not in used_answer_nums and zhuji_answer_block_looks_memorization(current):
            result[body_num] = (body_num, current or [])
            used_answer_nums.add(body_num)
            continue

        candidates = [
            (index, answer_num, block)
            for index, (answer_num, block) in enumerate(ordered)
            if answer_num not in used_answer_nums and zhuji_answer_block_looks_memorization(block)
        ]
        if candidates:
            candidates.sort(key=lambda item: (abs(item[1] - body_num), item[0]))
            _index, answer_num, block = candidates[0]
            result[body_num] = (answer_num, renumber_answer_block(block, body_num))
            used_answer_nums.add(answer_num)
        elif current:
            result[body_num] = (body_num, current)
            used_answer_nums.add(body_num)
    return result


def build_answer_piece_for_body(body_piece: wenyan.Piece, answer_block: list[str] | None) -> wenyan.Piece:
    if not answer_block:
        return wenyan.Piece(body_piece.volume, body_piece.num, body_piece.title, "", [body_piece.title])
    by_num = answer_blocks_by_question(answer_block)
    ordered = ordered_answer_blocks_by_question(answer_block)
    used_answer_nums: set[int] = set()
    lines = [body_piece.title]
    splitter = zuhe_split_sections_with_context if body_piece.volume in ZHUJI_MODULE2_GROUPS else wenyan.split_sections_with_context
    for _module, section_title, section_lines in splitter(body_piece):
        nums = body_section_question_nums(section_lines)
        if not nums:
            if is_zhuji_unnumbered_choice_section(section_lines) and zhuji_answer_block_is_bare_choice(answer_block):
                section_match = re.match(r"^([一二三四五六七八九十]+)、", plain(section_title))
                if section_match:
                    lines.append(f"{section_match.group(1)}、")
                elif section_title != "导入":
                    lines.append(section_title)
                lines.extend(answer_block)
            continue
        section_match = re.match(r"^([一二三四五六七八九十]+)、", plain(section_title))
        if section_match:
            lines.append(f"{section_match.group(1)}、")
        elif section_title != "导入":
            lines.append(section_title)
        resolved: dict[int, tuple[int, list[str]]] = {}
        if body_piece.volume in ZHUJI_MODULE2_GROUPS and "文言一词多义" in plain(section_title):
            resolved = zhuji_resolve_vocab_answer_blocks(section_lines, by_num, ordered, used_answer_nums)
        elif body_piece.volume in ZHUJI_MODULE2_GROUPS and "名篇名句默写" in plain(section_title):
            resolved = zhuji_resolve_memorization_answer_blocks(section_lines, by_num, ordered, used_answer_nums)
        for num in nums:
            if num in resolved:
                _answer_num, block = resolved[num]
                lines.extend(block)
                continue
            block = by_num.get(num, [])
            if block:
                used_answer_nums.add(num)
                lines.extend(block)
    return wenyan.Piece(body_piece.volume, body_piece.num, body_piece.title, "", lines)


def parse_answer_pieces(md_path: Path, body_pieces: list[wenyan.Piece]) -> list[wenyan.Piece]:
    blocks = parse_answer_piece_blocks(md_path)
    return [build_answer_piece_for_body(piece, blocks.get((piece.volume, piece.num))) for piece in body_pieces]


def section_summary_name(section_title: str) -> str:
    text = plain(section_title)
    text = re.sub(r"^[一二三四五六七八九十]+、\s*", "", text)
    text = re.sub(r"[「『【\[(（][^」』】\])）]*分[」』】\])）]", "", text)
    text = text.strip()
    return text


def piece_summary(piece: wenyan.Piece) -> str:
    names: list[str] = []
    for _module, section_title, _section_lines in wenyan.split_sections_with_context(piece):
        name = section_summary_name(section_title)
        if name in ZUHE_SUMMARY_ORDER and name not in names:
            names.append(name)
    names.sort(key=lambda item: ZUHE_SUMMARY_ORDER.index(item))
    return "+".join(names)


def format_zuhe_section_heading(line: str) -> str:
    text = plain(line)
    if not is_zuhe_major_section_title(line):
        return line

    def repl(match: re.Match[str]) -> str:
        return (
            "「"
            f'<span class="font-times">{match.group(1)}</span>'
            '<span class="font-song">分</span>'
            "」"
        )

    return re.sub(r"「(\d+)分」", repl, line)


def mark_zuhe_body_lines(lines: list[str]) -> list[str]:
    return [format_zuhe_section_heading(line) for line in lines]


def zuhe_split_analysis(text: str) -> tuple[str, list[str]]:
    marker_pattern = "|".join(re.escape(marker) for marker in ZUHE_ANALYSIS_MARKERS)
    first = re.search(marker_pattern, text)
    if not first:
        return text, []
    answer_part = text[: first.start()].strip()
    tail = text[first.start() :].strip()
    starts = list(re.finditer(marker_pattern, tail))
    analyses: list[str] = []
    for index, match in enumerate(starts):
        end = starts[index + 1].start() if index + 1 < len(starts) else len(tail)
        item = tail[match.start() : end].strip()
        if item:
            analyses.append(item)
    return answer_part, analyses


def zuhe_strip_answer_label(text: str) -> str:
    text = wenyan.strip_inline_tags(text).strip()
    text = re.sub(r"^\d+[\.、]\s*", "", text).strip()
    text = re.sub(r"^（?[一二三四五六七八九十]+）\s*", "", text).strip()
    text = re.sub(r"^（?\d+）\s*", "", text).strip()
    return text.strip("_＿ \t")


def zuhe_split_non_unique_answer_tail(answer_part: str) -> tuple[str, str | None]:
    if "答案不唯一" not in answer_part:
        return answer_part, None
    marker_index = answer_part.find("答案不唯一")
    for match in re.finditer(r"[（(]", answer_part[:marker_index]):
        index = match.start()
        if index > 0 and answer_part[index - 1] in "［[":
            continue
        if re.match(r"[（(]\d+[）)]", answer_part[index : index + 5]):
            continue
        if not re.search(r"[_＿]{2,}", answer_part[:index]):
            continue
        main = answer_part[:index].strip()
        tail = wenyan.strip_inline_tags(re.sub(r"[_＿]{2,}", " ", answer_part[index:])).strip()
        tail = re.sub(r"\s+", " ", tail)
        return main, tail or None
    return answer_part, None


def zuhe_answer_text_to_values(text: str) -> list[str]:
    answer_part, _analyses = zuhe_split_analysis(text.strip())
    answer_part = answer_part.strip()
    if not answer_part:
        return []

    plain_answer = wenyan.strip_inline_tags(answer_part).strip()
    has_blanks = bool(re.search(r"[_＿]{2,}", plain_answer))
    circled_re = re.compile(f"[{re.escape(wenyan.CIRCLED_NUMBERS)}]")

    # Open-answer summaries use circled numbers as visible answer bullets.
    if not has_blanks and circled_re.search(plain_answer):
        return [zuhe_strip_answer_label(plain_answer)]

    if has_blanks:
        values: list[str] = []
        for raw_part in re.split(r"[_＿]{2,}", plain_answer):
            part = zuhe_strip_answer_label(raw_part)
            if not part:
                continue
            values.append(part)
        return values

    return [value for value in [zuhe_strip_answer_label(plain_answer)] if value]


def zuhe_append_answer_item_text(item: wenyan.AnswerItem, text: str, last_kind: str | None) -> str | None:
    answer_part, analyses = zuhe_split_analysis(text.strip())
    if answer_part:
        answer_part, optional_tail = zuhe_split_non_unique_answer_tail(answer_part)
        if answer_part:
            item.answers.extend(zuhe_answer_text_to_values(answer_part))
            last_kind = "answer"
        if optional_tail:
            item.analyses.append(optional_tail)
            last_kind = "analysis"
    if analyses:
        item.analyses.extend(analyses)
        last_kind = "analysis"
    elif not answer_part and last_kind == "analysis" and item.analyses:
        item.analyses[-1] = f"{item.analyses[-1]}{text.strip()}"
    return last_kind


def move_zuhe_non_unique_tail_from_previous_answer(item: wenyan.AnswerItem, continuation: str) -> bool:
    if not item.answers:
        return False
    previous = item.answers[-1]
    match = re.search(r"[（(](?!\d+[）)])", previous)
    if not match:
        return False
    main = previous[: match.start()].strip()
    tail = f"{previous[match.start():]}　{continuation.strip()}"
    tail = wenyan.strip_inline_tags(re.sub(r"[_＿]{2,}", " ", tail)).strip()
    tail = re.sub(r"\s+", " ", tail)
    if main:
        item.answers[-1] = main
    else:
        item.answers.pop()
    if tail:
        item.analyses.append(tail)
    return True


def zuhe_extract_answer_items(answer_lines: list[str] | None) -> dict[int, wenyan.AnswerItem]:
    if not answer_lines:
        return {}
    items: dict[int, wenyan.AnswerItem] = {}
    current: wenyan.AnswerItem | None = None
    last_kind: str | None = None
    skipping_reference_tail = False
    for raw_line in answer_lines:
        line = raw_line.strip()
        if not line or line in wenyan.MODULE_TITLES or line in wenyan.VOLUMES or wenyan.is_section_marker(line):
            skipping_reference_tail = False
            last_kind = None
            continue
        match = wenyan.answer_item_match(line)
        if match:
            skipping_reference_tail = False
            current = wenyan.AnswerItem(answers=[], analyses=[])
            items[int(match.group(1))] = current
            last_kind = zuhe_append_answer_item_text(current, match.group(2), None)
            continue
        if skipping_reference_tail:
            continue
        if current is not None and any(marker in line for marker in ZUHE_ANALYSIS_MARKERS):
            last_kind = zuhe_append_answer_item_text(current, line, last_kind)
            continue
        if wenyan.looks_like_reference_translation_tail(line, current):
            skipping_reference_tail = True
            continue
        if current is not None:
            if last_kind == "answer" and "答案不唯一" in line and move_zuhe_non_unique_tail_from_previous_answer(current, line):
                last_kind = "analysis"
                continue
            answer_part, analyses = zuhe_split_analysis(line)
            if analyses:
                last_kind = zuhe_append_answer_item_text(current, line, last_kind)
            elif last_kind == "analysis" and current.analyses:
                current.analyses[-1] = f"{current.analyses[-1]}{line}"
            elif answer_part:
                last_kind = zuhe_append_answer_item_text(current, line, last_kind)
    return items


def zuhe_split_long_answer_text(answer: str, max_units: float = ZUHE_LONG_ANSWER_TEXT_UNITS) -> list[str]:
    text = re.sub(r"\s+", "", answer).strip()
    if not text:
        return []
    pieces: list[str] = []
    break_chars = "；;。！？!?，,、）)]"
    while text:
        if wenyan.text_units(text) <= max_units:
            pieces.append(text)
            break
        units = 0.0
        hard_cut = 0
        breakpoints: list[tuple[int, float]] = []
        for index, ch in enumerate(text):
            units += 1.0 if ord(ch) > 127 else 0.55
            if ch in break_chars:
                breakpoints.append((index + 1, units))
            if units > max_units:
                hard_cut = index
                break
        candidates = [
            pos
            for pos, unit_count in breakpoints
            if unit_count >= max_units - ZUHE_LONG_ANSWER_TRAILING_UNITS
        ]
        cut = candidates[-1] if candidates else hard_cut
        if cut <= 0:
            cut = 1
        pieces.append(text[:cut].strip())
        text = text[cut:].strip()
    if len(pieces) >= 2 and wenyan.text_units(pieces[-1]) <= ZUHE_LONG_ANSWER_TRAILING_UNITS:
        previous_units = wenyan.text_units(pieces[-2])
        tail_units = wenyan.text_units(pieces[-1])
        if previous_units + tail_units <= ZUHE_LONG_ANSWER_UNDERLINE_UNITS:
            pieces[-2] = f"{pieces[-2]}{pieces[-1]}"
            pieces.pop()
    return [piece for piece in pieces if piece]


def zuhe_normalize_analysis_text(text: str) -> str:
    if "答案不唯一" in text:
        # Optional-answer tails need readable separation between alternative
        # paired verses; do not squash the spaces inserted while parsing blanks.
        return re.sub(r"\s+", " ", text).strip()
    return re.sub(r"\s+", "", text).strip()


def zuhe_split_analysis_text(text: str, max_units: float = ZUHE_ANALYSIS_FRAGMENT_UNITS) -> list[str]:
    text = zuhe_normalize_analysis_text(text)
    if not text or wenyan.text_units(text) <= max_units:
        return [text] if text else []
    fragments: list[str] = []
    while text:
        units = 0.0
        last_break = -1
        cut = len(text)
        for index, ch in enumerate(text):
            weight = 1.0 if ord(ch) > 127 else 0.55
            if units + weight > max_units:
                cut = index
                break
            units += weight
            if units >= max_units * 0.70 and ch in "。！？!?；;":
                last_break = index + 1
        if cut < len(text) and last_break > 0:
            cut = last_break
        if cut <= 0:
            cut = 1
        fragments.append(text[:cut].strip())
        text = text[cut:].strip()
    return [fragment for fragment in fragments if fragment]


def append_zuhe_analysis_lines(styled: list[wenyan.StyledLine], analysis: str) -> None:
    analysis = style_zuhe_analysis_markers(zuhe_normalize_analysis_text(analysis))
    if not analysis:
        return
    if wenyan.estimate_line_count(analysis) <= ZUHE_ANALYSIS_SINGLE_BLOCK_MAX_LINES:
        styled.append(wenyan.StyledLine(analysis, "analysis_block"))
        return
    fragments = zuhe_split_analysis_text(analysis)
    for index, fragment in enumerate(fragments):
        meta = {"force_new_chunk": True} if index > 0 else None
        styled.append(wenyan.StyledLine(fragment, "analysis_block", meta))


def zuhe_fixed_answer_line(answer: str = "") -> str:
    answer = answer.strip()
    leading = "　" if answer else ""
    answer_units = wenyan.text_units(f"{leading}{answer}")
    padding_units = max(
        int(ZUHE_LONG_ANSWER_TRAILING_UNITS),
        int(math.ceil(ZUHE_LONG_ANSWER_UNDERLINE_UNITS - answer_units)),
    )
    # Use underlined non-breaking spaces in the same text box instead of a
    # separate line shape, so long-answer blanks remain editable like normal
    # fill-in blanks while the answer overlay can still be cropped tightly.
    padding = "\u00a0" * (padding_units * 2)
    return wenyan.answer_span(f"{leading}{answer}{padding}", underline=True)


def zuhe_fixed_long_answer_lines(
    answer: str,
    *,
    min_lines: int = 1,
    max_units: float = ZUHE_LONG_ANSWER_TEXT_UNITS,
) -> list[wenyan.StyledLine]:
    pieces = zuhe_split_long_answer_text(answer, max_units=max_units)
    line_count = len(pieces) if pieces else max(1, min_lines)
    return [
        wenyan.StyledLine(zuhe_fixed_answer_line(pieces[index] if index < len(pieces) else ""), "long_answer")
        for index in range(line_count)
    ]


def is_zuhe_blank_only_line(line: str) -> bool:
    return bool(re.fullmatch(r"[_＿]{3,}", plain(line)))


def remove_zuhe_open_answer_blank_lines(lines: list[str]) -> list[str]:
    return [line for line in lines if not is_zuhe_blank_only_line(line)]


def zuhe_open_answer_line_count(lines: list[str]) -> int:
    return sum(1 for line in lines if is_zuhe_blank_only_line(line))


def is_zuhe_revision_prompt(lines: list[str]) -> bool:
    text = "\n".join(plain(line) for line in lines)
    return "表述不当" in text and "序号" in text and "可修改为" in text


def is_zuhe_translation_prompt(lines: list[str]) -> bool:
    text = "\n".join(plain(line) for line in lines[:2])
    return "翻译成现代汉语" in text


def is_zuhe_circled_answer_blank_line(line: str) -> bool:
    return bool(re.fullmatch(rf"[{re.escape(wenyan.CIRCLED_NUMBERS)}]\s*[_＿]{{2,}}", plain(line)))


def is_zuhe_circled_answer_row(line: str) -> bool:
    text = plain(line).replace(" ", "").replace("　", "")
    return bool(re.fullmatch(rf"(?:[{re.escape(wenyan.CIRCLED_NUMBERS)}][＿_]+)+", text))


def is_zuhe_qa_completion_prompt(lines: list[str]) -> bool:
    text = "\n".join(plain(line) for line in lines)
    return "问答文段" in text and "补写" in text and any(is_zuhe_circled_answer_blank_line(line) for line in lines)


def zuhe_clean_qa_answer_value(answer: str) -> str:
    answer = strip_zuhe_score_note(wenyan.strip_inline_tags(answer)).strip()
    answer = re.sub(rf"^[{re.escape(wenyan.CIRCLED_NUMBERS)}]\s*", "", answer).strip()
    answer = re.sub(r"[_＿]+", " ", answer)
    return re.sub(r"\s+", " ", answer).strip()


def zuhe_qa_answer_values(answers: list[str]) -> list[str]:
    joined = "".join(wenyan.strip_inline_tags(answer) for answer in answers)
    circled_matches = list(re.finditer(rf"[{re.escape(wenyan.CIRCLED_NUMBERS)}]", joined))
    if len(circled_matches) >= 2:
        values: list[str] = []
        for index, match in enumerate(circled_matches):
            end = circled_matches[index + 1].start() if index + 1 < len(circled_matches) else len(joined)
            value = zuhe_clean_qa_answer_value(joined[match.start() : end])
            if value:
                values.append(value)
        if values:
            return values
    return [value for value in (zuhe_clean_qa_answer_value(answer) for answer in answers) if value]


def fill_zuhe_qa_completion_answers(lines: list[str], answers: list[str]) -> tuple[list[str], bool]:
    result: list[str] = []
    answer_values = zuhe_qa_answer_values(answers)
    answer_index = 0
    filled = False
    for line in lines:
        if is_zuhe_circled_answer_blank_line(line) and answer_index < len(answer_values):
            prefix = plain(line)[0]
            answer = answer_values[answer_index]
            result.append(f"{prefix}{wenyan.answer_span(answer, underline=True, pad_underline=True)}")
            answer_index += 1
            filled = True
            continue
        result.append(line)
    return result, filled


def fill_zuhe_circled_answer_row(lines: list[str], answers: list[str]) -> tuple[list[str], bool]:
    answer_values = zuhe_circled_row_answer_values(answers)
    if not answer_values:
        return lines, False
    result: list[str] = []
    filled = False
    for line in lines:
        if is_zuhe_circled_answer_row(line):
            output = line
            answer_index = 0

            def repl(match: re.Match[str]) -> str:
                nonlocal answer_index, filled
                if answer_index >= len(answer_values):
                    return match.group(0)
                value = answer_values[answer_index]
                answer_index += 1
                filled = True
                return f"{match.group(1)}{wenyan.answer_span(value, underline=True, pad_underline=True)}"

            output = re.sub(rf"([{re.escape(wenyan.CIRCLED_NUMBERS)}])\s*[_＿]+", repl, output)
            result.append(output)
            continue
        result.append(line)
    return result, filled


def zuhe_split_revision_answer(answer: str, max_units: float = ZUHE_REVISION_INLINE_UNITS) -> list[str]:
    text = re.sub(r"\s+", "", wenyan.strip_inline_tags(answer)).strip()
    if not text:
        return []
    if wenyan.text_units(text) <= max_units:
        return [text]
    pieces: list[str] = []
    break_chars = "；;。！？!?，,、）)]/"
    while text:
        if wenyan.text_units(text) <= max_units:
            pieces.append(text)
            break
        units = 0.0
        hard_cut = 0
        breakpoints: list[tuple[int, float]] = []
        for index, ch in enumerate(text):
            weight = 1.0 if ord(ch) > 127 else 0.55
            units += weight
            if ch in "（(" and units >= max_units * 0.55:
                breakpoints.append((index, units - weight))
            elif ch in break_chars and units >= max_units * 0.55:
                breakpoints.append((index + 1, units))
            if units > max_units:
                hard_cut = index
                break
        candidates = [
            pos
            for pos, unit_count in breakpoints
            if max_units * 0.55 <= unit_count <= max_units
        ]
        cut = candidates[-1] if candidates else hard_cut
        while cut > 1 and cut < len(text) and text[cut - 1].isdigit() and text[cut].isdigit():
            cut -= 1
        if cut <= 0:
            cut = 1
        pieces.append(text[:cut].strip())
        text = text[cut:].strip()
    return pieces or [text]


def zuhe_circled_row_answer_values(answers: list[str]) -> list[str]:
    joined = "".join(wenyan.strip_inline_tags(answer) for answer in answers)
    joined = strip_zuhe_score_note(joined)
    values: list[str] = []
    circled_matches = list(re.finditer(rf"[{re.escape(wenyan.CIRCLED_NUMBERS)}]", joined))
    for index, match in enumerate(circled_matches):
        end = circled_matches[index + 1].start() if index + 1 < len(circled_matches) else len(joined)
        chunk = joined[match.start() : end]
        chunk = re.sub(rf"^[{re.escape(wenyan.CIRCLED_NUMBERS)}]\s*", "", chunk).strip()
        chunk = re.sub(r"[_＿]+", " ", chunk)
        chunk = re.sub(r"\s+", " ", chunk).strip()
        if chunk:
            values.append(chunk)
    if values:
        return values
    return [zuhe_strip_answer_label(wenyan.strip_inline_tags(answer)) for answer in answers if zuhe_strip_answer_label(answer)]


def zuhe_interleaved_translation_lines(
    group_lines: list[str],
    answers: list[str],
    section_title: str,
) -> tuple[list[wenyan.StyledLine], bool]:
    styled: list[wenyan.StyledLine] = []
    answer_index = 0
    for line in group_lines:
        if is_zuhe_blank_only_line(line):
            continue
        styled.extend(zuhe_styled_body_lines([line], section_title))
        if re.match(r"^（\d+）", plain(line)) and answer_index < len(answers):
            styled.extend(zuhe_fixed_long_answer_lines(answers[answer_index], min_lines=1, max_units=24.0))
            answer_index += 1
    return styled, answer_index > 0


def strip_zuhe_score_note(answer: str) -> str:
    return re.sub(r"[（(][^（）()]{0,12}(?:每处|共|分)[^（）()]*[）)]\s*$", "", answer).strip()


def fill_zuhe_revision_answers(lines: list[str], answers: list[str]) -> tuple[list[str], bool]:
    if not answers:
        return lines, False
    answer_index = 0
    filled = 0
    result: list[str] = []
    skip_blank_only = False
    for line in lines:
        if skip_blank_only and is_zuhe_blank_only_line(line):
            skip_blank_only = False
            continue
        skip_blank_only = False
        if "序号" in plain(line) and "可修改为" in plain(line) and answer_index < len(answers):
            continuation_lines: list[str] = []
            number = strip_zuhe_score_note(answers[answer_index])
            answer_index += 1
            line, _idx, number_filled = wenyan.fill_plain_underline_blanks(line, [number], 0)
            filled += number_filled
            if answer_index < len(answers):
                revision = strip_zuhe_score_note(answers[answer_index])
                answer_index += 1
                revision_parts = zuhe_split_revision_answer(revision)

                def repl(match: re.Match[str]) -> str:
                    first_part = revision_parts[0] if revision_parts else revision
                    return wenyan.answer_span(first_part, underline=True, pad_underline=True)

                line, count = re.subn(r"(?<=可修改为：)[_＿]{2,}", repl, line, count=1)
                filled += count
                continuation_lines = [
                    f"{'　' * 13}{wenyan.answer_span(continuation, underline=True, pad_underline=True)}"
                    for continuation in revision_parts[1:]
                ]
                skip_blank_only = count > 0
            result.append(line)
            result.extend(continuation_lines)
            continue
        result.append(line)
    return result, filled > 0


def zuhe_extract_replacement_pairs(text: str) -> list[str]:
    plain_text = wenyan.strip_inline_tags(text).strip()
    if "改为" not in plain_text:
        return []
    pairs: list[str] = []
    quote = r"[“”\"'‘’]?"
    pattern = re.compile(
        rf"{quote}([^“”\"'‘’；;，,。！？\s]{{1,24}}){quote}\s*改为\s*"
        rf"{quote}([^“”\"'‘’；;，,。！？\s]{{1,24}}){quote}"
    )
    for match in pattern.finditer(plain_text):
        pairs.extend([match.group(1), match.group(2)])
    return pairs


def is_zuhe_replacement_prompt(lines: list[str]) -> bool:
    text = "\n".join(plain(line) for line in lines)
    return "改为" in text and "可修改为" not in text and "序号" not in text and "答案不唯一" not in text


def fill_zuhe_replacement_answers(lines: list[str], answers: list[str]) -> tuple[list[str], bool]:
    answer_values = zuhe_extract_replacement_pairs("；".join(answers))
    if not answer_values:
        return lines, False
    result: list[str] = []
    answer_index = 0
    filled = False
    for line in lines:
        plain_line = plain(line)
        compact = re.sub(r"[_＿\s　]+", "", plain_line)
        if compact and re.fullmatch(r"(改为)+", compact):
            pair_lines: list[str] = []
            while answer_index + 1 < len(answer_values):
                pair_lines.append(
                    f"{wenyan.answer_span(answer_values[answer_index], underline=True)}"
                    f"改为{wenyan.answer_span(answer_values[answer_index + 1], underline=True)}"
                )
                answer_index += 2
            if pair_lines:
                result.append("　　".join(pair_lines))
                filled = True
            else:
                result.append(line)
            continue
        result.append(line)

    if filled:
        return result, True
    return wenyan.fill_lines_with_answer_values(lines, answer_values)


def zuhe_paragraph_properties(role: str) -> dict[str, str]:
    attrs = wenyan._zuhe_original_paragraph_properties(role)  # type: ignore[attr-defined]
    if role == "long_answer":
        attrs["indent"] = "0"
        attrs["marL"] = "0"
        attrs["algn"] = "ctr"
    if role in {"vocab_table_header", "vocab_table_answer", "vocab_table_header_compact", "vocab_table_answer_compact"}:
        attrs["indent"] = "0"
        attrs["marL"] = "0"
        attrs["algn"] = "ctr"
    if role in {
        "vocab_table_title",
        "vocab_table_prompt",
        "vocab_table_meaning",
        "vocab_table_cell",
        "vocab_table_meaning_compact",
        "vocab_table_cell_compact",
    }:
        attrs["indent"] = "0"
        attrs["marL"] = "0"
    if role == "material_reading":
        attrs["indent"] = "0"
        attrs["marL"] = "0"
        attrs["algn"] = "just"
    return attrs


def zuhe_role_style(role: str) -> tuple[str, float, str, bool]:
    if role == "vocab_table_title":
        return "黑体", 22, "000000", True
    if role == "vocab_table_prompt":
        return "宋体", 19, "000000", False
    if role == "vocab_table_meaning":
        return "宋体", 18, "000000", False
    if role == "vocab_table_meaning_compact":
        return "宋体", 16, "000000", False
    if role == "vocab_table_header":
        return "黑体", 18, "000000", True
    if role == "vocab_table_header_compact":
        return "黑体", 16, "000000", True
    if role == "vocab_table_cell":
        return "宋体", 16, "000000", False
    if role == "vocab_table_cell_compact":
        return "宋体", 14, "000000", False
    if role == "vocab_table_answer":
        return "宋体", 18, "FF0000", True
    if role == "vocab_table_answer_compact":
        return "宋体", 16, "FF0000", True
    if role == "material_reading":
        return "KaiTi", 24, "000000", False
    if role in {"source_note", "annotation"}:
        return "仿宋", 24, "000000", False
    return wenyan._zuhe_original_role_style(role)  # type: ignore[attr-defined]


def append_zuhe_textbox(
    root,
    *,
    x: float,
    y: float,
    w: float,
    h: float,
    paragraphs: list[tuple[str, list[tuple]]],
    name: str,
    rels_root=None,
    slide_number: int | None = None,
) -> int:
    shape_id = wenyan.next_shape_id(root)
    sp = wenyan.ET.Element(wenyan.qn(wenyan.P_NS, "sp"))

    nv_sp_pr = wenyan.ET.SubElement(sp, wenyan.qn(wenyan.P_NS, "nvSpPr"))
    c_nv_pr = wenyan.ET.SubElement(
        nv_sp_pr,
        wenyan.qn(wenyan.P_NS, "cNvPr"),
        {"id": str(shape_id), "name": f"{name} {shape_id}"},
    )
    if rels_root is not None and slide_number is not None:
        rel_id = wenyan.add_slide_jump_relationship(rels_root, slide_number)
        wenyan.ET.SubElement(
            c_nv_pr,
            wenyan.qn(wenyan.A_NS, "hlinkClick"),
            {wenyan.qn(wenyan.R_NS, "id"): rel_id, "action": "ppaction://hlinksldjump"},
        )
    wenyan.ET.SubElement(nv_sp_pr, wenyan.qn(wenyan.P_NS, "cNvSpPr"), {"txBox": "1"})
    wenyan.ET.SubElement(nv_sp_pr, wenyan.qn(wenyan.P_NS, "nvPr"))

    sp_pr = wenyan.ET.SubElement(sp, wenyan.qn(wenyan.P_NS, "spPr"))
    xfrm = wenyan.ET.SubElement(sp_pr, wenyan.qn(wenyan.A_NS, "xfrm"))
    wenyan.ET.SubElement(xfrm, wenyan.qn(wenyan.A_NS, "off"), {"x": str(wenyan.inch(x)), "y": str(wenyan.inch(y))})
    wenyan.ET.SubElement(xfrm, wenyan.qn(wenyan.A_NS, "ext"), {"cx": str(wenyan.inch(w)), "cy": str(wenyan.inch(h))})
    geom = wenyan.ET.SubElement(sp_pr, wenyan.qn(wenyan.A_NS, "prstGeom"), {"prst": "rect"})
    wenyan.ET.SubElement(geom, wenyan.qn(wenyan.A_NS, "avLst"))
    wenyan.ET.SubElement(sp_pr, wenyan.qn(wenyan.A_NS, "noFill"))
    ln = wenyan.ET.SubElement(sp_pr, wenyan.qn(wenyan.A_NS, "ln"))
    wenyan.ET.SubElement(ln, wenyan.qn(wenyan.A_NS, "noFill"))

    tx_body = wenyan.ET.SubElement(sp, wenyan.qn(wenyan.P_NS, "txBody"))
    body_pr = wenyan.ET.SubElement(
        tx_body,
        wenyan.qn(wenyan.A_NS, "bodyPr"),
        {"wrap": "square", "anchor": "t", "lIns": "0", "tIns": "0", "rIns": "0", "bIns": "0"},
    )
    wenyan.ET.SubElement(body_pr, wenyan.qn(wenyan.A_NS, "spAutoFit"))
    wenyan.ET.SubElement(tx_body, wenyan.qn(wenyan.A_NS, "lstStyle"))

    for align, runs in paragraphs:
        p = wenyan.ET.SubElement(tx_body, wenyan.qn(wenyan.A_NS, "p"))
        p_pr = wenyan.ET.SubElement(p, wenyan.qn(wenyan.A_NS, "pPr"), {"algn": align})
        line_spacing = wenyan.ET.SubElement(p_pr, wenyan.qn(wenyan.A_NS, "lnSpc"))
        wenyan.ET.SubElement(line_spacing, wenyan.qn(wenyan.A_NS, "spcPct"), {"val": "120000"})
        last_size = 2000
        for run in runs:
            text, font, size, color, bold, *extra = run
            italic = bool(extra[0]) if extra else False
            wenyan.append_run(p, text, font=font, size=size, color=color, bold=bold)
            if italic:
                r = p.findall(wenyan.qn(wenyan.A_NS, "r"))[-1]
                r_pr = r.find(wenyan.qn(wenyan.A_NS, "rPr"))
                if r_pr is not None:
                    r_pr.attrib["i"] = "1"
            last_size = int(size * 100)
        wenyan.ET.SubElement(p, wenyan.qn(wenyan.A_NS, "endParaRPr"), {"lang": "zh-CN", "sz": str(last_size)})

    wenyan.append_to_sp_tree(root, sp)
    return shape_id


def remove_zuhe_named_shapes(root, name_prefix: str) -> None:
    sp_tree = root.find(f".//{{{wenyan.P_NS}}}spTree")
    if sp_tree is None:
        return
    for child in list(sp_tree):
        if child.tag != wenyan.qn(wenyan.P_NS, "sp"):
            continue
        c_nv_pr = child.find(f".//{{{wenyan.P_NS}}}cNvPr")
        if c_nv_pr is not None and c_nv_pr.attrib.get("name", "").startswith(name_prefix):
            sp_tree.remove(child)


def append_zuhe_cover_group_label(root, group: str, start: int, end: int) -> None:
    remove_zuhe_named_shapes(root, "组合加分练封面组别")
    append_zuhe_textbox(
        root,
        x=4.00,
        y=4.78,
        w=8.45,
        h=0.96,
        name="组合加分练封面组别",
        paragraphs=[
            ("ctr", [(group, "Songti SC Black", 30, "000000", True)]),
            ("ctr", [(f"第{start}-{end}练", "Songti SC Black", 30, "000000", True)]),
        ],
    )


def add_zuhe_cover_content(root, volume: str, start: int, end: int) -> None:
    original = getattr(wenyan, "_zuhe_original_add_cover_content", None)
    if original is not None:
        original(root, volume, start, end)
    append_zuhe_cover_group_label(root, volume, start, end)


def add_zhuji_cover_content(root, volume: str, start: int, end: int) -> None:
    """Module cover is already designed in the 筑基 template; avoid extra overlays."""
    for sp in wenyan.text_shapes(root):
        text = wenyan.get_shape_text(sp).strip()
        if text == "模块一　语用专项40练":
            wenyan.set_shape_text(sp, "模块一　语用专项40练")


def italicize_zuhe_title_volume_label(root) -> None:
    for sp in root.findall(f".//{{{wenyan.P_NS}}}sp"):
        c_nv_pr = sp.find(f".//{{{wenyan.P_NS}}}cNvPr")
        if c_nv_pr is None or not c_nv_pr.attrib.get("name", "").startswith("篇目标题册别"):
            continue
        for r_pr in sp.findall(f".//{{{wenyan.A_NS}}}rPr"):
            r_pr.attrib["i"] = "1"


def zuhe_directory_y_positions(count: int) -> list[float]:
    if count <= 3:
        return [1.28, 2.66, 4.04][:count]
    start = 0.94
    step = min(1.16, 4.75 / max(1, count - 1))
    return [start + index * step for index in range(count)]


def append_zuhe_directory_item(root, rels_root, piece: wenyan.Piece, slide_number: int, y: float) -> None:
    summary = piece_summary(piece)
    append_zuhe_textbox(
        root,
        x=3.86,
        y=y,
        w=6.4,
        h=0.58,
        name=f"组合加分练目录标题 {piece.num}",
        rels_root=rels_root,
        slide_number=slide_number,
        paragraphs=[
            (
                "l",
                [
                    ("·", "黑体", 32, "000000", True),
                    ("第 ", "黑体", 32, "000000", True),
                    (str(piece.num), "黑体", 32, "000000", True),
                    (" 练", "黑体", 32, "000000", True),
                ],
            )
        ],
    )
    if summary:
        append_zuhe_textbox(
            root,
            x=4.19,
            y=y + 0.67,
            w=8.1,
            h=0.42,
            name=f"组合加分练目录概要 {piece.num}",
            rels_root=rels_root,
            slide_number=slide_number,
            paragraphs=[("l", [(summary, "黑体", 20, "000000", False)])],
        )


def add_zuhe_directory_content(root, rels_root, pieces: list[wenyan.Piece], title_slide_numbers: dict[int, int]) -> None:
    for sp in wenyan.text_shapes(root):
        if wenyan.get_shape_text(sp).strip() and "各篇目点击即可直达" not in wenyan.get_shape_text(sp):
            wenyan.set_shape_text(sp, "")

    for piece, y in zip(pieces, zuhe_directory_y_positions(len(pieces))):
        append_zuhe_directory_item(root, rels_root, piece, title_slide_numbers[piece.num], y)


def add_zhuji_directory_content(root, rels_root, pieces: list[wenyan.Piece], title_slide_numbers: dict[int, int]) -> None:
    for sp in wenyan.text_shapes(root):
        text = wenyan.get_shape_text(sp).strip()
        if text and "各篇目点击即可直达" not in text and "课件目录" not in text:
            wenyan.set_shape_text(sp, "")

    left_x = 3.15
    right_x = 7.10
    start_y = 1.04
    step = 0.48
    for index, piece in enumerate(pieces):
        column = 0 if index < 10 else 1
        row = index if index < 10 else index - 10
        x = left_x if column == 0 else right_x
        y = start_y + row * step
        append_zuhe_textbox(
            root,
            x=x,
            y=y,
            w=3.65,
            h=0.34,
            name=f"筑基目录 {piece.num}",
            rels_root=rels_root,
            slide_number=title_slide_numbers[piece.num],
            paragraphs=[("l", [(piece.title, "黑体", 18, "000000", True)])],
        )


def zhuji_modules_for_pieces(pieces: list[wenyan.Piece]) -> list[str]:
    modules: list[str] = []
    for piece in pieces:
        if piece.volume in ZHUJI_MODULE_TITLES and piece.volume not in modules:
            modules.append(piece.volume)
    return modules


def zhuji_directory_groups_for_pieces(pieces: list[wenyan.Piece]) -> list[tuple[str, list[wenyan.Piece]]]:
    modules = set(zhuji_modules_for_pieces(pieces))
    result: list[tuple[str, list[wenyan.Piece]]] = []
    for title, group_modules in ZHUJI_TOP_DIRECTORY_GROUPS:
        if not any(module in modules for module in group_modules):
            continue
        grouped = [piece for piece in pieces if piece.volume in group_modules]
        if grouped:
            result.append((title, grouped))
    for module in zhuji_modules_for_pieces(pieces):
        if any(module in group_modules for _title, group_modules in ZHUJI_TOP_DIRECTORY_GROUPS):
            continue
        grouped = [piece for piece in pieces if piece.volume == module]
        if grouped:
            result.append((module, grouped))
    return result


def zhuji_detail_directory_groups_for_pieces(pieces: list[wenyan.Piece]) -> list[tuple[str, list[wenyan.Piece]]]:
    result: list[tuple[str, list[wenyan.Piece]]] = []
    for module in zhuji_modules_for_pieces(pieces):
        grouped = [piece for piece in pieces if piece.volume == module]
        if grouped:
            result.append((module, grouped))
    return result


def zhuji_directory_uses_vertical_layout(grouped: list[wenyan.Piece]) -> bool:
    if len(grouped) <= 3:
        return False
    max_title_units = max((wenyan.text_units(piece.title) for piece in grouped), default=0.0)
    return len(grouped) >= 4 or max_title_units > 16.0


def add_zhuji_top_directory_content(
    root,
    rels_root,
    pieces: list[wenyan.Piece],
    group_directory_slide_numbers: dict[str, int],
) -> None:
    for sp in wenyan.text_shapes(root):
        text = wenyan.get_shape_text(sp).strip()
        if text and "课件目录" not in text:
            wenyan.set_shape_text(sp, "")

    y_positions = [1.35, 2.45, 3.55, 4.65]
    for y, (label, _grouped) in zip(y_positions, zhuji_directory_groups_for_pieces(pieces)):
        slide_number = group_directory_slide_numbers.get(label)
        if slide_number is None:
            continue
        append_zuhe_textbox(
            root,
            x=4.40,
            y=y,
            w=5.90,
            h=0.55,
            name=f"筑基一级目录 {label}",
            rels_root=rels_root,
            slide_number=slide_number,
            paragraphs=[("l", [("· ", "黑体", 32, "000000", True), (label, "黑体", 32, "000000", True)])],
        )


def add_zhuji_group_directory_content(
    root,
    rels_root,
    group_title: str,
    group_pieces: list[wenyan.Piece],
    title_slide_numbers: dict[int, int],
) -> None:
    for sp in wenyan.text_shapes(root):
        text = wenyan.get_shape_text(sp).strip()
        if text and "各篇目点击即可直达" not in text and "课件目录" not in text:
            wenyan.set_shape_text(sp, "")

    append_zuhe_textbox(
        root,
        x=4.70,
        y=0.55,
        w=5.10,
        h=0.60,
        name=f"筑基三级目录标题 {group_title}",
        paragraphs=[("ctr", [(f"「{group_title}」", "黑体", 40, "E77818", False)])],
    )

    columns = 1 if zhuji_directory_uses_vertical_layout(group_pieces) else (2 if len(group_pieces) > 3 else 1)
    left_x = 3.55 if columns == 2 else 3.40
    right_x = 7.85
    start_y = 1.82
    step = 0.74 if columns == 1 and len(group_pieces) >= 5 else (1.10 if len(group_pieces) <= 6 else 0.82)
    split_at = math.ceil(len(group_pieces) / columns)
    for index, piece in enumerate(group_pieces):
        column = 0 if index < split_at else 1
        row = index if column == 0 else index - split_at
        x = left_x if column == 0 else right_x
        y = start_y + row * step
        append_zuhe_textbox(
            root,
            x=x,
            y=y,
            w=8.75 if columns == 1 else 4.00,
            h=0.50 if columns == 1 else 0.48,
            name=f"筑基三级目录 {piece.num}",
            rels_root=rels_root,
            slide_number=title_slide_numbers[piece.num],
            paragraphs=[("l", [("· ", "黑体", 28, "000000", True), (piece.title, "黑体", 28, "000000", True)])],
        )


def zhuji_title_parts(piece: wenyan.Piece) -> tuple[str, str]:
    text = plain(piece.title)
    match = re.match(r"^(第\d+练)(?:[:：](.*))?$", text)
    if not match:
        return text, ""
    return match.group(1), (match.group(2) or "").strip()


def add_zuhe_piece_title(root, piece: wenyan.Piece) -> None:
    wenyan.clear_text_shapes(root)
    summary = piece_summary(piece)
    append_zuhe_textbox(
        root,
        x=1.02,
        y=2.48,
        w=4.05,
        h=1.23,
        name=f"组合加分练标题 {piece.num}",
        paragraphs=[
            (
                "ctr",
                [
                    ("第 ", "方正小标宋_GBK", 80, "000000", True),
                    (str(piece.num), "方正小标宋_GBK", 80, ZUHE_TEAL, True),
                    (" 练", "方正小标宋_GBK", 80, "000000", True),
                ],
            )
        ],
    )
    wenyan.append_rect(root, x=5.38, y=2.40, w=0.035, h=1.48, fill_color="000000", line_color=None, name="组合加分练标题竖线")
    if summary:
        append_zuhe_textbox(
            root,
            x=6.05,
            y=3.05,
            w=5.9,
            h=0.45,
            name=f"组合加分练标题概要 {piece.num}",
            paragraphs=[("l", [(summary, "黑体", 20, "000000", False)])],
        )
    wenyan.append_title_volume_label(root, piece.volume)
    italicize_zuhe_title_volume_label(root)


def add_zhuji_piece_title(root, piece: wenyan.Piece) -> None:
    wenyan.clear_text_shapes(root)
    title, subtitle = zhuji_title_parts(piece)
    left_x = 2.25
    left_w = 4.05
    right_x = 6.55
    right_w = 4.80
    center_y = 3.00
    append_zuhe_textbox(
        root,
        x=left_x,
        y=center_y - 0.72,
        w=left_w,
        h=1.15,
        name=f"筑基标题 {piece.num}",
        paragraphs=[
            (
                "ctr",
                [
                    ("第 ", "方正小标宋_GBK", 80, "000000", True),
                    (str(piece.num), "方正小标宋_GBK", 80, "EE822F", True),
                    (" 练", "方正小标宋_GBK", 80, "000000", True),
                ],
            )
        ],
    )
    if subtitle:
        append_zuhe_textbox(
            root,
            x=right_x,
            y=center_y - 0.31,
            w=right_w,
            h=0.62,
            name=f"筑基标题题型 {piece.num}",
            paragraphs=[("ctr", [(subtitle, "黑体", 36, "000000", False)])],
        )
    wenyan.append_rect(root, x=6.35, y=center_y - 0.70, w=0.035, h=1.40, fill_color="000000", line_color=None, name="筑基标题竖线")
    # The title template carries an inherited "模块一" footer label; the user
    # wants no extra text in the lower-left corner on title pages.
    wenyan.append_rect(root, x=0.00, y=6.25, w=3.20, h=0.85, fill_color="E77818", line_color=None, name="筑基标题底部遮罩")


def is_zuhe_question_start(text: str) -> bool:
    return bool(re.match(r"^\d+[\.．、]\s*", text) or re.match(r"^[A-D][\.．、]\s*", text))


def split_zuhe_memorization_answer_line(line: str) -> list[str]:
    """Keep 名篇名句默写 inline answers away from long prompt text.

    These questions often put the actual blanks inside Chinese quotes after a
    long prompt. If the hidden answer text participates in the same paragraph,
    PowerPoint may wrap it back over the prompt. Splitting at the quote keeps
    the answer-bearing text on its own line while preserving the original order.
    """

    if "answer-fill" not in line:
        return [line]
    answer_index = line.find("answer-fill")
    open_quote = line.rfind("“", 0, answer_index)
    if open_quote < 0:
        return [line]
    prefix = line[:open_quote].rstrip()
    if wenyan.text_units(plain(prefix)) < 18:
        return [line]
    close_quote = line.find("”", answer_index)
    if close_quote < 0:
        return [prefix, line[open_quote:].lstrip()]

    quoted = line[open_quote : close_quote + 1].strip()
    suffix = line[close_quote + 1 :].strip()
    if suffix and wenyan.text_units(plain(quoted + suffix)) <= 42:
        return [prefix, f"{quoted}{suffix}"]
    result = [prefix, quoted]
    if suffix:
        result.append(suffix)
    return result


def split_zuhe_memorization_answer_lines(lines: list[str], section_title: str) -> list[str]:
    if "名篇名句默写" not in plain(section_title):
        return lines
    result: list[str] = []
    for line in lines:
        result.extend(split_zuhe_memorization_answer_line(line))
    return result


def zhuji_circled_answer_values(answers: list[str]) -> list[str]:
    text = "".join(wenyan.strip_inline_tags(answer) for answer in answers)
    return re.findall(f"[{re.escape(wenyan.CIRCLED_NUMBERS)}]", text)


def is_zhuji_vocab_table_item(group_lines: list[str]) -> bool:
    plain_lines = [plain(line) for line in group_lines if plain(line)]
    return (
        len(plain_lines) >= 5
        and bool(re.match(r"^\d+[\.．、]\s*\S+", plain_lines[0]))
        and any(line.startswith("常见义项：") for line in plain_lines)
        and "例句义项" in plain_lines
        and "高考链接" in plain_lines
    )


def zhuji_vocab_prompt_lines(lines: list[str]) -> list[str]:
    return [line for line in lines if "选择合适的义项" in plain(line)]


def zhuji_vocab_intro_lines(lines: list[str]) -> list[str]:
    intro: list[str] = []
    for line in lines:
        if is_zuhe_major_section_title(line):
            intro.append(line)
    intro.extend(zhuji_vocab_prompt_lines(lines))
    return intro


def zhuji_is_vocab_section_boundary(line: str) -> bool:
    text = plain(line).strip()
    return text in {
        "一",
        "二",
        "三",
        "四",
        "语言文字运用",
        "文言一词多义",
        "名篇名句默写",
    } or is_zuhe_major_section_title(line)


def zhuji_split_vocab_intro_tail(lines: list[str]) -> tuple[list[str], list[str]]:
    """Peel the one-word-multiple-meaning prompt out of the previous question."""

    for index, line in enumerate(lines):
        if "选择合适的义项" not in plain(line):
            continue
        start = index
        while start > 0 and zhuji_is_vocab_section_boundary(lines[start - 1]):
            start -= 1
        return lines[:start], zhuji_vocab_prompt_lines(lines[start:])
    return lines, []


def zhuji_split_vocab_group_tail(lines: list[str]) -> tuple[list[str], list[str]]:
    """Keep the following module heading from being swallowed by the last table item."""

    seen_link = False
    link_rows = 0
    for index, line in enumerate(lines):
        text = plain(line).strip()
        if text == "高考链接":
            seen_link = True
            link_rows = 0
            continue
        if not seen_link:
            continue
        if link_rows >= 1 and zhuji_is_vocab_section_boundary(line):
            return lines[:index], lines[index:]
        if text:
            link_rows += 1
    return lines, []


def zhuji_parse_vocab_table_item(
    group_lines: list[str],
    item: wenyan.AnswerItem,
    intro_lines: list[str] | None = None,
) -> dict[str, object] | None:
    table_lines, _tail_lines = zhuji_split_vocab_group_tail(group_lines)
    cleaned = [line.strip() for line in table_lines if line.strip()]
    plain_lines = [plain(line) for line in cleaned]
    if not is_zhuji_vocab_table_item(cleaned):
        return None
    try:
        meaning_index = next(index for index, line in enumerate(plain_lines) if line.startswith("常见义项："))
        example_index = plain_lines.index("例句义项")
        link_index = plain_lines.index("高考链接")
    except ValueError:
        return None
    examples = cleaned[example_index + 1 : link_index]
    links = cleaned[link_index + 1 :]
    if not examples and not links:
        return None
    answers = zhuji_circled_answer_values(item.answers)
    row_count = len(examples) + len(links)
    if len(answers) < row_count:
        answers.extend([""] * (row_count - len(answers)))
    return {
        "intro_lines": intro_lines or [],
        "title": cleaned[0],
        "meanings": cleaned[meaning_index],
        "examples": examples,
        "links": links,
        "answers": answers[:row_count],
    }


def zhuji_vocab_table_section_lines(
    segments: list[tuple[str, int | None, list[str]]],
    answer_items: dict[int, wenyan.AnswerItem],
) -> list[wenyan.StyledLine] | None:
    groups = [(num, group_lines) for kind, num, group_lines in segments if kind == "group" and num is not None]
    if not groups or not all(is_zhuji_vocab_table_item(group_lines) for _num, group_lines in groups):
        return None

    intro_lines: list[str] = []
    for kind, _num, segment_lines in segments:
        if kind == "leading":
            intro_lines.extend(zhuji_vocab_intro_lines(segment_lines))

    result: list[wenyan.StyledLine] = []
    for num, group_lines in groups:
        item = answer_items.get(num)
        if item is None:
            return None
        table = zhuji_parse_vocab_table_item(group_lines, item, intro_lines if not result else None)
        if table is None:
            return None
        result.append(wenyan.StyledLine("", "vocab_table", table))
    return result


def zuhe_split_body_numbered_segments(
    lines: list[str],
) -> list[tuple[str, int | None, list[str]]]:
    """Split numbered questions without swallowing the next minor reading section.

    组合加分练 often keeps “（一）语言文字运用Ⅰ / （二）语言文字运用Ⅱ”
    under one major section.  The generic splitter only stops at the next
    numbered question, so an open-answer question immediately before “（二）”
    can accidentally absorb the next article and place its answer after it.
    """

    segments: list[tuple[str, int | None, list[str]]] = []
    pending: list[str] = []
    current_num: int | None = None
    current_lines: list[str] = []

    def flush_pending() -> None:
        nonlocal pending
        if pending:
            segments.append(("leading", None, pending))
            pending = []

    def flush_group() -> None:
        nonlocal current_num, current_lines
        if current_num is not None:
            segments.append(("group", current_num, current_lines))
            current_num = None
            current_lines = []

    for line in lines:
        match = wenyan.body_item_match(line)
        if match:
            flush_group()
            flush_pending()
            current_num = int(match.group(1))
            current_lines = [line]
            continue

        if current_num is not None and is_zuhe_minor_section_title(line):
            flush_group()
            pending = [line]
            continue

        if current_num is None:
            pending.append(line)
        else:
            current_lines.append(line)

    flush_group()
    flush_pending()
    return segments


def zuhe_split_body_numbered_groups(lines: list[str]) -> tuple[list[str], list[tuple[int, list[str]]]]:
    leading: list[str] = []
    groups: list[tuple[int, list[str]]] = []
    for kind, num, segment_lines in zuhe_split_body_numbered_segments(lines):
        if kind == "leading" and not groups:
            leading.extend(segment_lines)
        elif kind == "group" and num is not None:
            groups.append((num, segment_lines))
    return leading, groups


def split_zuhe_reference_blocks(answer_lines: list[str] | None) -> tuple[list[str] | None, list[list[str]]]:
    if not answer_lines:
        return answer_lines, []
    cleaned: list[str] = []
    reference_blocks: list[list[str]] = []
    index = 0
    while index < len(answer_lines):
        line = answer_lines[index]
        if is_zuhe_reference_tif_line(line):
            block = [line]
            index += 1
            while index < len(answer_lines):
                next_line = answer_lines[index]
                if (
                    wenyan.answer_item_match(next_line)
                    or wenyan.is_section_marker(next_line)
                    or next_line in wenyan.MODULE_TITLES
                    or next_line in wenyan.VOLUMES
                ):
                    break
                block.append(next_line)
                index += 1
            reference_blocks.append(block)
            continue
        cleaned.append(line)
        index += 1
    return cleaned, reference_blocks


def append_zuhe_reference_block(styled: list[wenyan.StyledLine], block: list[str]) -> None:
    if not block:
        return
    styled.append(wenyan.StyledLine("", "page_break"))
    for index, line in enumerate(block):
        if is_zuhe_reference_tif_line(line):
            styled.append(wenyan.StyledLine(line, "body"))
        elif is_zuhe_source_note_plain(plain(line)):
            styled.append(wenyan.StyledLine(line, "source_note"))
        elif plain(line).startswith("【注】"):
            styled.append(wenyan.StyledLine(line, "annotation"))
        elif index == 1 and plain(line).endswith("："):
            styled.append(wenyan.StyledLine(line, "subheading"))
        elif zuhe_material_label_parts(line):
            material_line, material_role = style_zuhe_material_label(line)
            styled.append(wenyan.StyledLine(material_line, material_role))
        else:
            styled.append(wenyan.StyledLine(mark_zuhe_wave_blue(line), "extended_reading"))


def zuhe_styled_body_lines(lines: list[str], section_title: str = "") -> list[wenyan.StyledLine]:
    styled: list[wenyan.StyledLine] = []
    reading = wenyan.is_reading_section(section_title)
    plain_lines = [wenyan.strip_inline_tags(line).strip() for line in lines]
    in_article = False
    for index, line in enumerate(lines):
        plain_text = plain_lines[index]
        if line in wenyan.MODULE_TITLES or line in wenyan.VOLUMES or wenyan.is_piece_boundary_noise_line(line):
            continue
        center_text = wenyan.canonical_center_heading(line)
        if center_text:
            styled.append(wenyan.StyledLine(center_text, "center_heading"))
            in_article = False
            continue
        if is_zuhe_major_section_title(line):
            styled.append(wenyan.StyledLine(line, "subheading"))
            in_article = False
            continue
        if is_zuhe_source_note_plain(plain_text):
            styled.append(wenyan.StyledLine(line, "source_note"))
            continue
        if plain_text.startswith("【注】"):
            styled.append(wenyan.StyledLine(line, "annotation"))
            continue
        if in_article and zuhe_material_label_parts(line):
            material_line, material_role = style_zuhe_material_label(line)
            styled.append(wenyan.StyledLine(material_line, material_role))
            continue
        if wenyan.is_subheading_line(lines, index):
            styled.append(wenyan.StyledLine(line, "subheading"))
            in_article = False
            continue
        if "阅读下面" in plain_text or "阅读下列" in plain_text:
            styled.append(wenyan.StyledLine(line, "prompt"))
            in_article = True
            continue
        if in_article and plain_text and not is_zuhe_question_start(plain_text):
            styled.append(wenyan.StyledLine(mark_zuhe_wave_blue(line), "extended_reading"))
            continue
        if is_zuhe_question_start(plain_text):
            in_article = False
        if reading:
            styled.append(wenyan.StyledLine(line, "reading"))
        else:
            styled.append(wenyan.StyledLine(line, "body"))
    return styled


def zuhe_build_styled_section(lines: list[str], answer_lines: list[str] | None) -> tuple[list[wenyan.StyledLine], bool]:
    section_title = lines[0] if lines else ""
    core_answer_lines, reference_blocks = split_zuhe_reference_blocks(answer_lines)
    if is_zhuji_unnumbered_choice_section(lines) and zhuji_answer_block_is_bare_choice(core_answer_lines):
        answers = wenyan.extract_ordered_answers(core_answer_lines)
        filled_lines, filled = wenyan.fill_lines_with_answer_values(lines, answers)
        styled = zuhe_styled_body_lines(filled_lines, section_title)
        if filled:
            for analysis in wenyan.extract_unitemized_analyses(core_answer_lines):
                append_zuhe_analysis_lines(styled, analysis)
        for reference_block in reference_blocks:
            append_zuhe_reference_block(styled, reference_block)
        return styled, filled or bool(reference_blocks)

    answer_items = zuhe_extract_answer_items(core_answer_lines)
    segments = zuhe_split_body_numbered_segments(lines)
    groups = [(num, segment_lines) for kind, num, segment_lines in segments if kind == "group" and num is not None]
    if answer_items and groups:
        vocab_table_lines = zhuji_vocab_table_section_lines(segments, answer_items)
        if vocab_table_lines is not None:
            return vocab_table_lines, True

        styled: list[wenyan.StyledLine] = []
        matched_any = False
        preamble_filled_once = False
        vocab_intro_lines: list[str] = []

        for kind, num, group_lines in segments:
            if kind == "leading":
                group_lines, intro_lines = zhuji_split_vocab_intro_tail(group_lines)
                if intro_lines:
                    vocab_intro_lines = intro_lines
                if not group_lines:
                    continue
                if not preamble_filled_once:
                    preamble_lines, preamble_filled = wenyan.fill_lines_with_answer_values(
                        group_lines,
                        wenyan.extract_preamble_answers(core_answer_lines),
                    )
                    preamble_styled = zuhe_styled_body_lines(preamble_lines, section_title)
                    styled.extend(preamble_styled)
                    matched_any = matched_any or preamble_filled
                    preamble_filled_once = True
                else:
                    styled.extend(zuhe_styled_body_lines(group_lines, section_title))
                continue
            if num is None:
                group_lines, intro_lines = zhuji_split_vocab_intro_tail(group_lines)
                if intro_lines:
                    vocab_intro_lines = intro_lines
                styled.extend(zuhe_styled_body_lines(group_lines, section_title))
                continue
            group_lines, intro_lines = zhuji_split_vocab_intro_tail(group_lines)
            if intro_lines:
                vocab_intro_lines = intro_lines
            if not group_lines:
                continue
            item = answer_items.get(num)
            if item is None:
                styled.extend(zuhe_styled_body_lines(group_lines, section_title))
                continue

            table_group_lines, table_tail_lines = zhuji_split_vocab_group_tail(group_lines)
            if is_zhuji_vocab_table_item(table_group_lines):
                table = zhuji_parse_vocab_table_item(
                    table_group_lines,
                    item,
                    vocab_intro_lines,
                )
                if table is not None:
                    styled.append(wenyan.StyledLine("", "vocab_table", table))
                    vocab_intro_lines = []
                    if table_tail_lines:
                        styled.extend(zuhe_styled_body_lines(table_tail_lines, section_title))
                    matched_any = True
                    continue

            if wenyan.is_segmentation_prompt(group_lines):
                question_line = group_lines[0] if group_lines else ""
                answer = wenyan.answer_letters("".join(item.answers))
                styled.append(wenyan.StyledLine(wenyan.strip_fillable_blanks_from_line(question_line), "segmentation_question"))
                styled.append(
                    wenyan.StyledLine(
                        wenyan.segmentation_sentence_line(group_lines),
                        "segmentation_text",
                        {"answer": answer},
                    )
                )
                for analysis in item.analyses:
                    styled.append(wenyan.StyledLine(style_zuhe_analysis_markers(analysis), "segmentation_analysis"))
                styled.append(wenyan.StyledLine("", "page_break"))
                matched_any = True
                continue

            has_open_answer_lines = any(is_zuhe_blank_only_line(line) for line in group_lines)
            if item.answers and is_zuhe_revision_prompt(group_lines):
                filled_group, group_filled = fill_zuhe_revision_answers(group_lines, item.answers)
                styled.extend(zuhe_styled_body_lines(filled_group, section_title))
                matched_any = matched_any or group_filled
            elif item.answers and is_zuhe_replacement_prompt(group_lines):
                filled_group, group_filled = fill_zuhe_replacement_answers(group_lines, item.answers)
                styled.extend(zuhe_styled_body_lines(filled_group, section_title))
                if item.analyses:
                    styled.append(wenyan.StyledLine("", "answer_spacer"))
                matched_any = matched_any or group_filled
            elif item.answers and is_zuhe_qa_completion_prompt(group_lines):
                filled_group, group_filled = fill_zuhe_qa_completion_answers(group_lines, item.answers)
                styled.extend(zuhe_styled_body_lines(filled_group, section_title))
                matched_any = matched_any or group_filled
            elif item.answers and any(is_zuhe_circled_answer_row(line) for line in group_lines):
                filled_group, group_filled = fill_zuhe_circled_answer_row(group_lines, item.answers)
                styled.extend(zuhe_styled_body_lines(filled_group, section_title))
                if item.analyses:
                    styled.append(wenyan.StyledLine("", "answer_spacer"))
                matched_any = matched_any or group_filled
            elif item.answers and is_zuhe_translation_prompt(group_lines):
                translation_styled, group_filled = zuhe_interleaved_translation_lines(
                    group_lines,
                    item.answers,
                    section_title,
                )
                styled.extend(translation_styled)
                if item.analyses:
                    styled.append(wenyan.StyledLine("", "answer_spacer"))
                matched_any = matched_any or group_filled
            elif item.answers and has_open_answer_lines:
                question_lines = remove_zuhe_open_answer_blank_lines(group_lines)
                styled.extend(zuhe_styled_body_lines(question_lines, section_title))
                styled.extend(
                    zuhe_fixed_long_answer_lines(
                        "；".join(item.answers),
                        min_lines=zuhe_open_answer_line_count(group_lines),
                    )
                )
                if item.analyses:
                    styled.append(wenyan.StyledLine("", "answer_spacer"))
                matched_any = True
            elif item.answers and wenyan.is_long_answer_prompt(section_title, group_lines):
                question_lines = wenyan.remove_answer_blank_lines(group_lines)
                styled.extend(zuhe_styled_body_lines(question_lines, section_title))
                styled.extend(zuhe_fixed_long_answer_lines("；".join(item.answers)))
                if item.analyses:
                    styled.append(wenyan.StyledLine("", "answer_spacer"))
                matched_any = True
            elif wenyan.section_has_fillable_blanks(group_lines):
                filled_group, group_filled = wenyan.fill_lines_with_answer_values(group_lines, item.answers)
                styled.extend(zuhe_styled_body_lines(filled_group, section_title))
                matched_any = matched_any or group_filled
            else:
                styled.extend(zuhe_styled_body_lines(group_lines, section_title))
                for answer in item.answers:
                    styled.append(wenyan.StyledLine(answer, "answer"))
                matched_any = matched_any or bool(item.answers)

            for analysis in item.analyses:
                append_zuhe_analysis_lines(styled, analysis)
                matched_any = True
        for reference_block in reference_blocks:
            append_zuhe_reference_block(styled, reference_block)
            matched_any = True
        return styled, matched_any

    if wenyan.section_has_fillable_blanks(lines):
        filled_lines, embedded_answers = wenyan.embed_answers_in_lines(lines, core_answer_lines)
        styled = zuhe_styled_body_lines(filled_lines, section_title)
        if embedded_answers:
            for analysis in wenyan.extract_unitemized_analyses(core_answer_lines):
                append_zuhe_analysis_lines(styled, analysis)
        for reference_block in reference_blocks:
            append_zuhe_reference_block(styled, reference_block)
        return styled, embedded_answers
    styled = zuhe_styled_body_lines(lines, section_title)
    matched_any = False
    for reference_block in reference_blocks:
        append_zuhe_reference_block(styled, reference_block)
        matched_any = True
    return styled, matched_any


def zuhe_question_numbers_in_lines(lines: list[str]) -> set[int]:
    _leading, groups = zuhe_split_body_numbered_groups(lines)
    return {num for num, _group_lines in groups}


def zuhe_supplement_answer_lines_by_question_number(
    *,
    answer_piece: wenyan.Piece | None,
    body_lines: list[str],
    answer_lines: list[str] | None,
) -> list[str] | None:
    """Patch section-boundary mismatches by borrowing same-number answers.

    Some 组合加分练 answer files split open questions into a bare section such
    as "四、", while the body keeps those questions inside the previous reading
    section.  Section-key matching then finds only part of the answers.  This
    helper keeps the normal section match first, then appends answer sections
    containing missing question numbers from the same practice.
    """

    if answer_piece is None:
        return answer_lines
    wanted = zuhe_question_numbers_in_lines(body_lines)
    if not wanted:
        return answer_lines
    combined = list(answer_lines or [])
    present = set(zuhe_extract_answer_items(combined).keys())
    missing = wanted - present
    if not missing:
        return combined if combined else answer_lines

    for _module, _title, candidate_lines in wenyan.split_sections_with_context(answer_piece):
        candidate_items = zuhe_extract_answer_items(candidate_lines)
        candidate_nums = set(candidate_items.keys())
        if not candidate_nums or not (candidate_nums & missing):
            continue
        combined.extend(candidate_lines)
        present.update(candidate_nums)
        missing = wanted - present
        if not missing:
            break
    return combined if combined else answer_lines


def zuhe_make_piece_slide_chunks(body_piece: wenyan.Piece, answer_piece: wenyan.Piece | None) -> list[list[wenyan.StyledLine]]:
    """Use a slightly fuller content area for 组合加分练 while leaving other books untouched."""

    answer_map = wenyan.answer_section_map(answer_piece)
    chunks: list[list[wenyan.StyledLine]] = []
    pending_module_header: str | None = None
    for module, title, lines in wenyan.split_sections_with_context(body_piece):
        if title in wenyan.MODULE_TITLES:
            pending_module_header = wenyan.canonical_module_title(title)
            lines = [line for line in lines[1:] if line.strip()]
            if not lines:
                continue
        answer_lines = answer_map.get(wenyan.section_context_key(module, title))
        if not answer_lines and module:
            module_answer_lines = answer_map.get(wenyan.section_context_key(module, module))
            if module_answer_lines and any(marker in line for marker in ZUHE_ANALYSIS_MARKERS for line in module_answer_lines):
                answer_lines = module_answer_lines
        answer_lines = zuhe_supplement_answer_lines_by_question_number(
            answer_piece=answer_piece,
            body_lines=lines,
            answer_lines=answer_lines,
        )
        styled, _matched_answers = zuhe_build_styled_section(lines, answer_lines)
        if pending_module_header:
            styled = [wenyan.StyledLine(pending_module_header, "module_header")] + styled
            pending_module_header = None
        chunk_max_units = 12 if body_piece.volume in ZHUJI_MODULE2_GROUPS else 13
        chunks.extend(zuhe_chunk_styled_lines(styled, max_units=chunk_max_units))
    return chunks


def zuhe_logical_line_blocks(lines: list[wenyan.StyledLine]) -> list[wenyan.LineBlock]:
    """Keep choices intact, but let long answer explanations paginate by paragraph."""

    blocks: list[wenyan.LineBlock] = []
    index = 0
    analysis_roles = {"analysis", "analysis_block"}
    while index < len(lines):
        line = lines[index]
        if line.role == "vocab_table":
            blocks.append(wenyan.LineBlock([line], "vocab_table"))
            index += 1
            continue
        if line.role in {"body", "reading", "extended_reading"} and "序号" in plain(line.text) and "可修改为" in plain(line.text):
            end = index + 1
            while (
                end < len(lines)
                and (lines[end].role or "body") in {"body", "reading", "extended_reading"}
                and "answer-fill" in lines[end].text
                and plain(lines[end].text).startswith("　")
            ):
                end += 1
            blocks.append(wenyan.LineBlock(lines[index:end], "revision_answer"))
            index = end
            continue
        if wenyan.is_numbered_question_line(line):
            end = index + 1
            while end < len(lines):
                next_line = lines[end]
                if next_line.role in {
                    "analysis",
                    "analysis_block",
                    "segmentation_analysis",
                    "heading",
                    "module_header",
                    "page_break",
                    "segmentation_question",
                }:
                    break
                if wenyan.is_numbered_question_line(next_line):
                    break
                end += 1
            body_lines = lines[index:end]
            if wenyan.is_choice_body_block(body_lines):
                blocks.append(wenyan.LineBlock(body_lines, "choice_body"))
                analysis_end = end
                while analysis_end < len(lines) and (lines[analysis_end].role or "body") in analysis_roles:
                    blocks.append(wenyan.LineBlock([lines[analysis_end]], "choice_analysis"))
                    analysis_end += 1
                index = analysis_end
                continue
        blocks.append(wenyan.LineBlock([line], "single"))
        index += 1
    return blocks


def zuhe_chunk_styled_lines(lines: list[wenyan.StyledLine], max_units: int = 12) -> list[list[wenyan.StyledLine]]:
    expanded_lines: list[wenyan.StyledLine] = []
    for line in lines:
        expanded_lines.extend(wenyan.split_oversized_styled_line(line))
    blocks = zuhe_logical_line_blocks(expanded_lines)

    chunks: list[list[wenyan.StyledLine]] = []
    current: list[wenyan.StyledLine] = []
    used = 0
    index = 0
    while index < len(blocks):
        block = blocks[index]
        line = block.lines[0]
        if line.meta and line.meta.get("force_new_chunk") and current:
            chunks.append(current)
            current = []
            used = 0
        if line.role == "module_header":
            if current:
                chunks.append(current)
            current = [line]
            used = 2
            index += 1
            continue
        if line.role == "page_break":
            if current:
                chunks.append(current)
                current = []
                used = 0
            index += 1
            continue
        if line.role == "segmentation_question" and current:
            chunks.append(current)
            current = []
            used = 0

        cost = wenyan.block_page_cost(block)
        if block.kind == "vocab_table":
            if current:
                chunks.append(current)
                current = []
                used = 0
            chunks.append(block.lines)
            index += 1
            continue
        starts_new_question = wenyan.is_numbered_question_line(line)
        analysis_heavy_lead = current and current[-1].role in {"analysis", "analysis_block"} and used >= 4 and starts_new_question
        follows_answer_explanation = (
            starts_new_question
            and any((current_line.role or "") in {"analysis", "analysis_block", "segmentation_analysis"} for current_line in current)
        )
        if analysis_heavy_lead or follows_answer_explanation:
            chunks.append(current)
            current = []
            used = 0

        if block.kind == "choice_body":
            next_block = blocks[index + 1] if index + 1 < len(blocks) and blocks[index + 1].kind == "choice_analysis" else None
            next_cost = wenyan.block_page_cost(next_block) if next_block else 0
            choice_analysis_limit = max(1, max_units - 1)
            preferred_cost = cost + next_cost if next_block and cost + next_cost <= choice_analysis_limit else cost
            if current and (used + preferred_cost > max_units or used + cost > max_units):
                chunks.append(current)
                current = []
                used = 0
            if cost > max_units:
                option_start = next(
                    (
                        line_index
                        for line_index, choice_line in enumerate(block.lines[1:], start=1)
                        if wenyan.is_choice_option_line(choice_line)
                    ),
                    -1,
                )
                if option_start > 1:
                    stem_lines = block.lines[:option_start]
                    option_lines = block.lines[option_start:]
                    if current:
                        chunks.append(current)
                        current = []
                        used = 0
                    chunks.append(stem_lines)
                    current = list(option_lines)
                    used = sum(wenyan.line_page_cost(choice_line) for choice_line in option_lines)
                    if next_block and used + next_cost <= max_units:
                        wenyan.append_block_to_current(current, next_block)
                        used += next_cost
                        index += 2
                    else:
                        index += 1
                    continue
            wenyan.append_block_to_current(current, block)
            used += cost
            if next_block:
                if used + next_cost > choice_analysis_limit and current:
                    chunks.append(current)
                    current = []
                    used = 0
                wenyan.append_block_to_current(current, next_block)
                used += next_cost
                index += 2
            else:
                index += 1
            continue

        keep_with_short_lead = (
            used <= 6
            and cost <= 4
            and used + cost <= max_units
            and line.role in {"body", "reading", "extended_reading"}
        )
        if current and used + cost > max_units and not keep_with_short_lead:
            if current[-1].role == "subheading":
                carry = current.pop()
                if current:
                    chunks.append(current)
                current = [carry]
                used = wenyan.line_page_cost(carry)
            else:
                chunks.append(current)
                current = []
                used = 0
        wenyan.append_block_to_current(current, block)
        used += cost
        index += 1
    if current:
        chunks.append(current)
    return chunks


def zhuji_vocab_row_height(text: str, col_width: float, *, compact: bool) -> float:
    capacity = max(16.0, col_width * (4.25 if compact else 3.75))
    line_count = max(1, math.ceil(wenyan.text_units(plain(text)) / capacity))
    return max(0.30 if compact else 0.36, (0.20 if compact else 0.24) * line_count + 0.13)


def append_zhuji_table_line(root, *, x: float, y: float, w: float, h: float, name: str) -> None:
    wenyan.append_rect(root, x=x, y=y, w=w, h=h, fill_color="000000", line_color=None, name=name)


def append_zhuji_vocab_cell(
    root,
    rels_root,
    media_assets: dict[str, bytes],
    tif_search_dirs: list[Path],
    *,
    x: float,
    y: float,
    w: float,
    h: float,
    text: str,
    role: str,
    name: str,
    answer_mode: str = "normal",
) -> int:
    line = wenyan.StyledLine(text, role)
    shape_id = wenyan.append_textbox(
        root,
        x=x,
        y=y,
        w=w,
        h=h,
        lines=[line],
        name=name,
        answer_mode=answer_mode,
    )
    wenyan.append_tif_overlays(
        root,
        rels_root,
        media_assets,
        x=x,
        y=y,
        w=w,
        lines=[line],
        tif_search_dirs=tif_search_dirs,
    )
    return shape_id


def append_zhuji_vocab_table(
    root,
    rels_root,
    *,
    media_assets: dict[str, bytes],
    tif_search_dirs: list[Path],
    content_box: tuple[float, float, float, float],
    table_line: wenyan.StyledLine,
    name: str,
) -> list[wenyan.AnimationTarget]:
    data = table_line.meta or {}
    intro_lines = [str(line) for line in data.get("intro_lines", []) if str(line).strip()]
    title = str(data.get("title", ""))
    meanings = str(data.get("meanings", ""))
    examples = [str(line) for line in data.get("examples", [])]
    links = [str(line) for line in data.get("links", [])]
    answers = [str(answer) for answer in data.get("answers", [])]

    base_x, base_y, base_w, base_h = content_box
    left = base_x + 0.35
    top = base_y + 0.03
    width = min(base_w - 0.70, 11.25)
    right_w = 1.45
    left_w = width - right_w
    bottom_limit = base_y + base_h - 0.08

    animations: list[wenyan.AnimationTarget] = []
    compact = len(examples) + len(links) >= 7
    prompt_role = "vocab_table_prompt"
    meaning_role = "vocab_table_meaning_compact" if compact else "vocab_table_meaning"
    header_role = "vocab_table_header_compact" if compact else "vocab_table_header"
    cell_role = "vocab_table_cell_compact" if compact else "vocab_table_cell"
    answer_role = "vocab_table_answer_compact" if compact else "vocab_table_answer"

    if intro_lines:
        for intro_index, intro_line in enumerate(intro_lines):
            intro_role = "subheading" if is_zuhe_major_section_title(intro_line) else prompt_role
            intro_height = 0.36 if intro_role == "subheading" else 0.34
            append_zhuji_vocab_cell(
                root,
                rels_root,
                media_assets,
                tif_search_dirs,
                x=left,
                y=top,
                w=width,
                h=intro_height,
                text=intro_line,
                role=intro_role,
                name=f"{name} vocab prompt {intro_index + 1}",
            )
            top += 0.42

    append_zhuji_vocab_cell(
        root,
        rels_root,
        media_assets,
        tif_search_dirs,
        x=left,
        y=top,
        w=width,
        h=0.34,
        text=title,
        role="vocab_table_title",
        name=f"{name} vocab title",
    )
    top += 0.42

    meaning_h = zhuji_vocab_row_height(meanings, width - 0.25, compact=compact)
    append_zhuji_vocab_cell(
        root,
        rels_root,
        media_assets,
        tif_search_dirs,
        x=left + 0.18,
        y=top,
        w=width - 0.30,
        h=meaning_h,
        text=meanings,
        role=meaning_role,
        name=f"{name} vocab meanings",
    )
    top += meaning_h + 0.10

    rows: list[tuple[str, str, bool]] = []
    for index, text in enumerate(examples):
        rows.append((text, answers[index] if index < len(answers) else "", False))
    if links:
        rows.append(("高考链接", "", True))
        offset = len(examples)
        for index, text in enumerate(links):
            rows.append((text, answers[offset + index] if offset + index < len(answers) else "", False))

    header_h = 0.34 if compact else 0.38
    row_heights = [
        (0.32 if compact else 0.36) if is_header else zhuji_vocab_row_height(text, left_w - 0.18, compact=compact)
        for text, _answer, is_header in rows
    ]
    table_h = header_h + sum(row_heights)
    max_table_h = max(2.0, bottom_limit - top)
    if table_h > max_table_h:
        scale = max_table_h / table_h
        header_h = max(0.28, header_h * scale)
        row_heights = [max(0.26 if compact else 0.30, height * scale) for height in row_heights]
        table_h = header_h + sum(row_heights)

    # Draw the editable table frame with PowerPoint shapes so the recovered
    # table remains adjustable after generation.
    wenyan.append_rect(root, x=left, y=top, w=width, h=table_h, fill_color=None, line_color="000000", name=f"{name} vocab table border")
    append_zhuji_table_line(root, x=left + left_w, y=top, w=0.01, h=table_h, name=f"{name} vocab table vertical")
    y_cursor = top + header_h
    append_zhuji_table_line(root, x=left, y=y_cursor, w=width, h=0.01, name=f"{name} vocab table line")
    for row_h in row_heights[:-1]:
        y_cursor += row_h
        append_zhuji_table_line(root, x=left, y=y_cursor, w=width, h=0.01, name=f"{name} vocab table line")

    cell_pad_x = 0.08
    cell_pad_y = 0.04
    append_zhuji_vocab_cell(
        root,
        rels_root,
        media_assets,
        tif_search_dirs,
        x=left,
        y=top + 0.03,
        w=left_w,
        h=header_h - 0.04,
        text="例句",
        role=header_role,
        name=f"{name} vocab header example",
    )
    append_zhuji_vocab_cell(
        root,
        rels_root,
        media_assets,
        tif_search_dirs,
        x=left + left_w,
        y=top + 0.03,
        w=right_w,
        h=header_h - 0.04,
        text="义项",
        role=header_role,
        name=f"{name} vocab header meaning",
    )

    row_top = top + header_h
    for index, ((text, answer, is_header), row_h) in enumerate(zip(rows, row_heights)):
        if is_header:
            append_zhuji_vocab_cell(
                root,
                rels_root,
                media_assets,
                tif_search_dirs,
                x=left,
                y=row_top + 0.02,
                w=left_w,
                h=row_h - 0.03,
                text=text,
                role=header_role,
                name=f"{name} vocab link header",
            )
        else:
            append_zhuji_vocab_cell(
                root,
                rels_root,
                media_assets,
                tif_search_dirs,
                x=left + cell_pad_x,
                y=row_top + cell_pad_y,
                w=left_w - cell_pad_x * 2,
                h=max(0.20, row_h - cell_pad_y * 2),
                text=text,
                role=cell_role,
                name=f"{name} vocab example {index + 1}",
            )
            answer_text = wenyan.answer_span(answer) if answer else ""
            cell_x = left + left_w
            cell_y = row_top - 0.05
            cell_h = row_h + 0.06
            append_zhuji_vocab_cell(
                root,
                rels_root,
                media_assets,
                tif_search_dirs,
                x=cell_x,
                y=cell_y,
                w=right_w,
                h=cell_h,
                text=answer_text,
                role=answer_role,
                name=f"{name} vocab answer {index + 1}",
                answer_mode="base",
            )
            animations.extend(
                wenyan.append_answer_overlays(
                    root,
                    x=cell_x,
                    y=cell_y,
                    w=right_w,
                    h=cell_h,
                    lines=[wenyan.StyledLine(answer_text, answer_role)],
                    name=f"{name} vocab answer {index + 1}",
                )
            )
        row_top += row_h
    return animations


def add_zuhe_content(
    root,
    rels_root,
    piece: wenyan.Piece,
    chunk: list[wenyan.StyledLine],
    *,
    media_assets: dict[str, bytes],
    tif_search_dirs: list[Path],
    content_box: tuple[float, float, float, float],
) -> None:
    wenyan.clear_text_shapes(root, keep_substrings=("能力",))
    module_title, content_lines = wenyan.split_module_header(chunk)
    if module_title:
        wenyan.add_module_header_image(root, rels_root, module_title)
    if any(line.role == "segmentation_text" for line in content_lines):
        wenyan.add_segmentation_content(root, rels_root, piece, content_lines, module_title=None)
        return
    if len(content_lines) == 1 and content_lines[0].role == "vocab_table":
        animations = append_zhuji_vocab_table(
            root,
            rels_root,
            media_assets=media_assets,
            tif_search_dirs=tif_search_dirs,
            content_box=content_box,
            table_line=content_lines[0],
            name=f"{piece.title} content",
        )
        wenyan.set_slide_animations(root, animations)
        return

    base_x, base_y, base_w, base_h = content_box
    y = max(base_y, 1.12) if module_title else base_y
    h = base_h - (y - base_y)
    wenyan.append_textbox(
        root,
        x=base_x,
        y=y,
        w=base_w,
        h=h,
        lines=content_lines,
        name=f"{piece.title} content",
        answer_mode="base",
    )
    wenyan.append_tif_overlays(
        root,
        rels_root,
        media_assets,
        x=base_x,
        y=y,
        w=base_w,
        lines=content_lines,
        tif_search_dirs=tif_search_dirs,
    )
    animations = wenyan.append_answer_overlays(
        root,
        x=base_x,
        y=y,
        w=base_w,
        h=h,
        lines=content_lines,
        name=f"{piece.title} content",
    )
    wenyan.set_slide_animations(root, animations)


def build_zhuji_deck(
    template: Path,
    output: Path,
    pieces: list[wenyan.Piece],
    answer_map: dict[tuple[str, int], wenyan.Piece],
    volume: str,
    start: int,
    end: int,
    *,
    tif_search_dirs: list[Path],
    multiline_answer_min_width_pt: float = 0.0,
) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(template, "r") as zin:
        slide_count_template = len([n for n in zin.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)])
        title_src, content_src, final_sources, content_box = wenyan.template_contract(slide_count_template)
        directory_src = 3
        top_directory_groups = zhuji_directory_groups_for_pieces(pieces)
        detail_directory_groups = zhuji_detail_directory_groups_for_pieces(pieces)
        directory_slide_count = 1 + len(detail_directory_groups)

        generated_slides: list[tuple[int, bytes, bytes]] = []
        extra_media_assets: dict[str, bytes] = {}
        piece_chunks: list[tuple[wenyan.Piece, list[list[wenyan.StyledLine]]]] = []
        title_slide_numbers: dict[int, int] = {}
        next_generated_slide = 3 + directory_slide_count
        for piece in pieces:
            chunks = wenyan.make_piece_slide_chunks(piece, answer_map.get(wenyan.piece_key(piece)))
            piece_chunks.append((piece, chunks))
            title_slide_numbers[piece.num] = next_generated_slide
            next_generated_slide += 1 + len(chunks)

        detail_directory_slide_numbers = {
            group_title: 4 + index for index, (group_title, _group_pieces) in enumerate(detail_directory_groups)
        }
        group_directory_slide_numbers: dict[str, int] = {}
        for group_title, group_pieces in top_directory_groups:
            first_piece = group_pieces[0]
            group_directory_slide_numbers[group_title] = detail_directory_slide_numbers[first_piece.volume]

        def add_from(src: int, edit_func=lambda root, rels_root: None) -> None:
            xml, rels = wenyan.make_slide_parts(zin, src, edit_func)
            generated_slides.append((src, xml, rels))

        add_from(1, lambda root, rels_root: add_zhuji_cover_content(root, volume, start, end))
        add_from(2)
        add_from(
            directory_src,
            lambda root, rels_root: add_zhuji_top_directory_content(
                root,
                rels_root,
                pieces,
                group_directory_slide_numbers,
            ),
        )
        for group_title, group_pieces in detail_directory_groups:
            add_from(
                directory_src,
                lambda root, rels_root, group_title=group_title, group_pieces=group_pieces: add_zhuji_group_directory_content(
                    root,
                    rels_root,
                    group_title,
                    group_pieces,
                    title_slide_numbers,
                ),
            )

        for piece, chunks in piece_chunks:
            add_from(title_src, lambda root, rels_root, piece=piece: add_zhuji_piece_title(root, piece))
            for chunk in chunks:
                add_from(
                    content_src,
                    lambda root, rels_root, piece=piece, chunk=chunk: add_zuhe_content(
                        root,
                        rels_root,
                        piece,
                        chunk,
                        media_assets=extra_media_assets,
                        tif_search_dirs=tif_search_dirs,
                        content_box=content_box,
                    ),
                )

        for src in final_sources:
            add_from(src)

        slide_total = len(generated_slides)
        content_types = wenyan.update_content_types(wenyan.parse_xml(zin.read("[Content_Types].xml")), slide_total)
        presentation_xml = wenyan.update_presentation_xml(wenyan.parse_xml(zin.read("ppt/presentation.xml")), slide_total)
        presentation_rels = wenyan.update_presentation_rels(
            wenyan.parse_xml(zin.read("ppt/_rels/presentation.xml.rels")),
            slide_total,
        )

        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            skip = {
                "[Content_Types].xml",
                "ppt/presentation.xml",
                "ppt/_rels/presentation.xml.rels",
            }
            for item in zin.infolist():
                name = item.filename
                if item.is_dir():
                    continue
                if name in skip:
                    continue
                if re.match(r"ppt/slides/slide\d+\.xml$", name):
                    continue
                if re.match(r"ppt/slides/_rels/slide\d+\.xml\.rels$", name):
                    continue
                zout.writestr(item, zin.read(name))
            written_media = {item.filename for item in zin.infolist()}
            for module in wenyan.MODULE_IMAGE_FILES:
                image_info = wenyan.module_image_package(module)
                if image_info is None:
                    continue
                image_path, package_path = image_info
                if package_path in written_media:
                    continue
                zout.writestr(package_path, image_path.read_bytes())
                written_media.add(package_path)
            for package_path, data in extra_media_assets.items():
                if package_path not in written_media:
                    zout.writestr(package_path, data)
                    written_media.add(package_path)
            zout.writestr("[Content_Types].xml", content_types)
            zout.writestr("ppt/presentation.xml", presentation_xml)
            zout.writestr("ppt/_rels/presentation.xml.rels", presentation_rels)
            for index, (_src, xml, rels) in enumerate(generated_slides, 1):
                zout.writestr(f"ppt/slides/slide{index}.xml", xml)
                zout.writestr(f"ppt/slides/_rels/slide{index}.xml.rels", rels)
    old_multiline_shift = getattr(wenyan, "ANSWER_MULTILINE_VERTICAL_SHIFT_PT", None)
    if old_multiline_shift is not None:
        wenyan.ANSWER_MULTILINE_VERTICAL_SHIFT_PT = ZHUJI_ANSWER_VERTICAL_SHIFT_PT
    try:
        wenyan.postprocess_tif_positions_with_powerpoint(
            output,
            multiline_answer_min_width_pt=multiline_answer_min_width_pt,
        )
    finally:
        if old_multiline_shift is not None:
            wenyan.ANSWER_MULTILINE_VERTICAL_SHIFT_PT = old_multiline_shift
    align_zhuji_long_answer_overlays(output)
    tighten_zhuji_answer_overlays(output)
    align_zhuji_choice_answer_overlays(output)


def tighten_zhuji_answer_overlays(pptx: Path) -> None:
    if sys.platform != "win32":
        return
    powershell = shutil.which("powershell") or shutil.which("pwsh")
    if not powershell:
        return
    script = r'''
$ErrorActionPreference = "Stop"
$pptx = (Resolve-Path -LiteralPath $env:PPTX_TO_TIGHTEN).Path
$pp = $null
$pres = $null
$changed = 0
try {
    $pp = New-Object -ComObject PowerPoint.Application
    $pres = $pp.Presentations.Open($pptx, $false, $false, $false)
    for ($si = 1; $si -le $pres.Slides.Count; $si++) {
        $slide = $pres.Slides.Item($si)
        for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
            $shape = $slide.Shapes.Item($shapeIndex)
            try {
                $name = [string]$shape.Name
                if ($name -notmatch "answer-adjusted") { continue }
                if ($shape.HasTextFrame -eq 0 -or $shape.TextFrame.HasText -eq 0) { continue }
                $text = [string]$shape.TextFrame.TextRange.Text
                $trimmedText = $text.Trim()
                if (-not $trimmedText) { continue }
                if ($text.Contains("【关键能力】") -or $text.Contains("【解题思路】") -or $text.Contains("【解决思路】") -or $text.Contains("【解析】")) {
                    continue
                }
                $hasManualBreak = $trimmedText.Contains("`r") -or $trimmedText.Contains("`n")
                if ((-not $hasManualBreak) -and $trimmedText.Length -gt 80) {
                    $shape.TextFrame.MarginLeft = 0
                    $shape.TextFrame.MarginRight = 0
                    $shape.TextFrame.MarginTop = 0
                    $shape.TextFrame.MarginBottom = 0
                    try { $shape.TextFrame.WordWrap = -1 } catch {}
                    $changed += 1
                }
                try { $shape.TextFrame.WordWrap = -1 } catch {}
                $tr = $shape.TextFrame.TextRange
                $boundLeft = [double]$tr.BoundLeft
                $boundTop = [double]$tr.BoundTop
                $boundWidth = [double]$tr.BoundWidth
                $boundHeight = [double]$tr.BoundHeight
                if ($boundWidth -le 1 -or $boundHeight -le 1) { continue }
                $shape.TextFrame.MarginLeft = 0
                $shape.TextFrame.MarginRight = 0
                $shape.TextFrame.MarginTop = 0
                $shape.TextFrame.MarginBottom = 0
                try { $shape.TextFrame.WordWrap = -1 } catch {}
                $targetLeft = [Math]::Max(0.0, $boundLeft - 1.5)
                $targetTop = [Math]::Max(0.0, $boundTop - 1.5)
                $targetWidth = [Math]::Max(18.0, $boundWidth + 4.0)
                try {
                    $slideWidth = [double]$pres.PageSetup.SlideWidth
                } catch {
                    $slideWidth = 960.0
                }
                $targetWidth = [Math]::Min($targetWidth, [Math]::Max(18.0, $slideWidth - $targetLeft - 8.0))
                $shape.Left = $targetLeft
                $shape.Top = $targetTop
                $shape.Width = $targetWidth
                $shape.Height = [Math]::Max(20.0, $boundHeight + 4.0)
                try {
                    $wrappedHeight = [double]$shape.TextFrame.TextRange.BoundHeight
                    if ($wrappedHeight -gt 1) {
                        $shape.Height = [Math]::Max([double]$shape.Height, $wrappedHeight + 6.0)
                    }
                } catch {}
                $changed += 1
            } catch {
                continue
            }
        }
    }
    if ($changed -gt 0) { $pres.Save() }
    Write-Output "ZHUJI_ANSWER_TIGHTENED=$changed"
}
finally {
    if ($pres -ne $null) {
        try { $pres.Close() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null } catch {}
    }
    if ($pp -ne $null) {
        try { $pp.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
'''
    env = os.environ.copy()
    env["PPTX_TO_TIGHTEN"] = str(pptx)
    try:
        result = subprocess.run(
            [powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            timeout=180,
        )
    except Exception as exc:
        print(f"提示：筑基答案文本框二次裁剪失败：{exc}", file=sys.stderr)
        return
    if result.returncode != 0:
        message = result.stderr.strip() or result.stdout.strip()
        print(f"提示：筑基答案文本框二次裁剪失败：{message}", file=sys.stderr)
        return
    output = result.stdout.strip()
    if output:
        print(f"筑基答案文本框二次裁剪：{pptx.name}，{output}")


def align_zhuji_choice_answer_overlays(pptx: Path) -> None:
    if sys.platform != "win32":
        return
    """Place single-letter choice answers in the stem parentheses.

    Generic text search can confuse the answer letter with option labels such
    as "D.".  This pass is deliberately choice-only: it requires A-D options in
    the same base text box and moves only one-letter answer overlays.
    """

    powershell = shutil.which("powershell") or shutil.which("pwsh")
    if not powershell:
        return
    script = r'''
$ErrorActionPreference = "Stop"
$pptx = (Resolve-Path -LiteralPath $env:PPTX_TO_ALIGN_CHOICE).Path
$pp = $null
$pres = $null
$changed = 0
function Get-BaseName([string]$shapeName) {
    $match = [regex]::Match($shapeName, "\s+answer\s+\d+\b")
    if (-not $match.Success) { return $null }
    return $shapeName.Substring(0, $match.Index).Trim().ToLowerInvariant()
}
function Get-BaseShape($slide, [string]$shapeName) {
    $baseName = Get-BaseName $shapeName
    if (-not $baseName) { return $null }
    for ($candidateIndex = 1; $candidateIndex -le $slide.Shapes.Count; $candidateIndex++) {
        try {
            $candidate = $slide.Shapes.Item($candidateIndex)
            $candidateName = ([string]$candidate.Name).ToLowerInvariant()
            if ($candidateName.Contains("answer")) { continue }
            if (-not $candidateName.StartsWith($baseName)) { continue }
            if ($candidate.HasTextFrame -eq 0 -or $candidate.TextFrame.HasText -eq 0) { continue }
            return $candidate
        } catch {
            continue
        }
    }
    return $null
}
function Has-ChoiceOptions([string]$text) {
    $seen = @{}
    foreach ($rawLine in @($text -split "[`r`n]+")) {
        $line = $rawLine.Trim(" ", "`t", "`r", "`n", [char]0x3000, [char]0x00A0)
        if ($line -match "^([A-D])[\.\uFF0E、]") {
            $seen[$matches[1]] = $true
        }
    }
    foreach ($letter in "A", "B", "C", "D") {
        if (-not $seen.ContainsKey($letter)) { return $false }
    }
    return $true
}
function Find-StemAnswerBox($baseShape, [string]$answerLetter) {
    $paragraphs = $baseShape.TextFrame.TextRange.Paragraphs()
    for ($pi = 1; $pi -le $paragraphs.Count; $pi++) {
        try {
            $paragraph = $baseShape.TextFrame.TextRange.Paragraphs($pi)
            $paragraphText = [string]$paragraph.Text
            if ($paragraphText -notmatch "^\s*\d+[\.\uFF0E、]") { continue }
            $pattern = "[（(]\s*" + [regex]::Escape($answerLetter) + "\s*[）)]"
            $match = [regex]::Match($paragraphText, $pattern)
            if (-not $match.Success) { continue }
            $relative = $paragraphText.IndexOf($answerLetter, $match.Index)
            if ($relative -lt 0) { continue }
            $range = $paragraph.Characters($relative + 1, 1)
            return [PSCustomObject]@{
                Left = [double]$range.BoundLeft
                Top = [double]$range.BoundTop
            }
        } catch {
            continue
        }
    }
    return $null
}
try {
    $pp = New-Object -ComObject PowerPoint.Application
    $pres = $pp.Presentations.Open($pptx, $false, $false, $false)
    for ($si = 1; $si -le $pres.Slides.Count; $si++) {
        $slide = $pres.Slides.Item($si)
        for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
            $shape = $slide.Shapes.Item($shapeIndex)
            try {
                $shapeName = [string]$shape.Name
                if ($shapeName -notmatch "answer-adjusted") { continue }
                if ($shape.HasTextFrame -eq 0 -or $shape.TextFrame.HasText -eq 0) { continue }
                $text = ([string]$shape.TextFrame.TextRange.Text).Trim()
                if ($text -notmatch "^[A-D]$") { continue }
                $baseShape = Get-BaseShape $slide $shapeName
                if ($baseShape -eq $null) { continue }
                $baseText = [string]$baseShape.TextFrame.TextRange.Text
                if (-not (Has-ChoiceOptions $baseText)) { continue }
                $target = Find-StemAnswerBox $baseShape $text
                if ($target -eq $null) { continue }
                $currentBoundLeft = [double]$shape.TextFrame.TextRange.BoundLeft
                $currentBoundTop = [double]$shape.TextFrame.TextRange.BoundTop
                $dx = [double]$target.Left - $currentBoundLeft
                $dy = [double]$target.Top - $currentBoundTop
                if ([Math]::Abs($dx) -lt 0.5 -and [Math]::Abs($dy) -lt 0.5) { continue }
                $shape.Left = [Math]::Max(0.0, [double]$shape.Left + $dx)
                $shape.Top = [Math]::Max(0.0, [double]$shape.Top + $dy)
                $changed += 1
            } catch {
                continue
            }
        }
    }
    if ($changed -gt 0) { $pres.Save() }
    Write-Output "ZHUJI_CHOICE_ANSWER_ALIGNED=$changed"
}
finally {
    if ($pres -ne $null) {
        try { $pres.Close() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null } catch {}
    }
    if ($pp -ne $null) {
        try { $pp.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
'''
    env = os.environ.copy()
    env["PPTX_TO_ALIGN_CHOICE"] = str(pptx)
    try:
        result = subprocess.run(
            [powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            timeout=180,
        )
    except Exception as exc:
        print(f"提示：筑基选择题答案校位失败：{exc}", file=sys.stderr)
        return
    if result.returncode != 0:
        message = result.stderr.strip() or result.stdout.strip()
        print(f"提示：筑基选择题答案校位失败：{message}", file=sys.stderr)
        return
    output = result.stdout.strip()
    if output:
        print(f"筑基选择题答案校位：{pptx.name}，{output}")


def align_zhuji_long_answer_overlays(pptx: Path) -> None:
    if sys.platform != "win32":
        return
    """Re-anchor long-answer overlays to the editable underline paragraphs.

    The generic cropper already gives the answer text boxes the right width and
    height.  For this book, long-answer blanks are rendered as underlined text
    paragraphs in the base content box, so this final pass only moves the red
    overlay boxes to the hidden answer glyph bounds in that base paragraph.
    """

    powershell = shutil.which("powershell") or shutil.which("pwsh")
    if not powershell:
        return
    script = r'''
$ErrorActionPreference = "Stop"
$pptx = (Resolve-Path -LiteralPath $env:PPTX_TO_ALIGN).Path
$pp = $null
$pres = $null
$changed = 0
function Get-BaseName([string]$shapeName) {
    $match = [regex]::Match($shapeName, "\s+answer\s+\d+\b")
    if (-not $match.Success) { return $null }
    return $shapeName.Substring(0, $match.Index).Trim().ToLowerInvariant()
}
function Normalize-Line([string]$text) {
    if ($null -eq $text) { return "" }
    return ($text -replace "[\s`u{3000}`u{00A0}_＿]+", "")
}
function Trim-AnswerLine([string]$text) {
    if ($null -eq $text) { return "" }
    return $text.Trim(" ", "`t", "`r", "`n", [char]0x3000, [char]0x00A0, "_", "＿")
}
function Is-AnalysisText([string]$text) {
    return $text.Contains("【关键能力】") -or
        $text.Contains("【解题思路】") -or
        $text.Contains("【解决思路】") -or
        $text.Contains("【解析】")
}
function Is-ChoiceAnswerOverlay($slide, [string]$shapeName, [string]$text) {
    $trimmed = $text.Trim()
    if ($trimmed -notmatch "^[A-D]$") { return $false }
    $baseShape = Get-BaseShape $slide $shapeName
    if ($baseShape -eq $null) { return $false }
    $paragraphs = $baseShape.TextFrame.TextRange.Paragraphs()
    $letters = @{}
    for ($pi = 1; $pi -le $paragraphs.Count; $pi++) {
        try {
            $line = ([string]$baseShape.TextFrame.TextRange.Paragraphs($pi).Text).Trim(" ", "`t", "`r", "`n", [char]0x3000, [char]0x00A0)
            if ($line -match "^([A-D])[\.\uFF0E、]") { $letters[$matches[1]] = $true }
        } catch {
            continue
        }
    }
    foreach ($letter in "A", "B", "C", "D") {
        if (-not $letters.ContainsKey($letter)) { return $false }
    }
    return $true
}
function Get-TextUnits([string]$text) {
    $units = 0.0
    foreach ($ch in $text.ToCharArray()) {
        if ([int][char]$ch -gt 127) { $units += 1.0 } else { $units += 0.55 }
    }
    return $units
}
function Get-BaseShape($slide, [string]$shapeName) {
    $baseName = Get-BaseName $shapeName
    if (-not $baseName) { return $null }
    for ($candidateIndex = 1; $candidateIndex -le $slide.Shapes.Count; $candidateIndex++) {
        try {
            $candidate = $slide.Shapes.Item($candidateIndex)
            $candidateName = ([string]$candidate.Name).ToLowerInvariant()
            if ($candidateName.Contains("answer")) { continue }
            if (-not $candidateName.StartsWith($baseName)) { continue }
            if ($candidate.HasTextFrame -eq 0 -or $candidate.TextFrame.HasText -eq 0) { continue }
            return $candidate
        } catch {
            continue
        }
    }
    return $null
}
function Find-AnswerLineBox($baseShape, [string]$lineText) {
    $needle = Trim-AnswerLine $lineText
    if (-not $needle) { return $null }
    $needleNorm = Normalize-Line $needle
    if (-not $needleNorm) { return $null }
    $paragraphs = $baseShape.TextFrame.TextRange.Paragraphs()
    for ($pi = 1; $pi -le $paragraphs.Count; $pi++) {
        try {
            $paragraph = $baseShape.TextFrame.TextRange.Paragraphs($pi)
            $paragraphText = [string]$paragraph.Text
            $paragraphTrimmed = Trim-AnswerLine $paragraphText
            if (-not $paragraphTrimmed) { continue }
            $paragraphNorm = Normalize-Line $paragraphTrimmed
            if (-not $paragraphNorm.StartsWith($needleNorm)) {
                $prefixLength = [Math]::Min(12, $needleNorm.Length)
                if ($prefixLength -lt 4 -or -not $paragraphNorm.StartsWith($needleNorm.Substring(0, $prefixLength))) {
                    continue
                }
            }
            $searchIndex = $paragraphText.IndexOf($needle)
            if ($searchIndex -lt 0) {
                $shortNeedle = $needle.Substring(0, [Math]::Min(12, $needle.Length))
                $searchIndex = $paragraphText.IndexOf($shortNeedle)
            }
            if ($searchIndex -lt 0) { continue }
            $length = [Math]::Min($needle.Length, [Math]::Max(1, $paragraphText.Length - $searchIndex))
            $range = $paragraph.Characters($searchIndex + 1, $length)
            return [PSCustomObject]@{
                Left = [double]$range.BoundLeft
                Top = [double]$range.BoundTop
                Right = [double]$range.BoundLeft + [double]$range.BoundWidth
                Bottom = [double]$range.BoundTop + [double]$range.BoundHeight
            }
        } catch {
            continue
        }
    }
    return $null
}
try {
    $pp = New-Object -ComObject PowerPoint.Application
    $pres = $pp.Presentations.Open($pptx, $false, $false, $false)
    for ($si = 1; $si -le $pres.Slides.Count; $si++) {
        $slide = $pres.Slides.Item($si)
        for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
            $shape = $slide.Shapes.Item($shapeIndex)
            try {
                $shapeName = [string]$shape.Name
                if ($shapeName -notmatch "answer-adjusted") { continue }
                if ($shape.HasTextFrame -eq 0 -or $shape.TextFrame.HasText -eq 0) { continue }
                $text = [string]$shape.TextFrame.TextRange.Text
                if (-not $text.Trim()) { continue }
                if (Is-AnalysisText $text) { continue }
                if (Is-ChoiceAnswerOverlay $slide $shapeName $text) { continue }
                $lineTexts = @($text -split "[`r`n]+" | ForEach-Object { Trim-AnswerLine $_ } | Where-Object { $_ })
                if ($lineTexts.Count -eq 0) { continue }
                $isLongAnswer = $lineTexts.Count -gt 1 -or (Get-TextUnits (($lineTexts -join "")) -gt 18.0)
                if (-not $isLongAnswer) { continue }
                $baseShape = Get-BaseShape $slide $shapeName
                if ($baseShape -eq $null) { continue }
                $boxes = @()
                foreach ($lineText in $lineTexts) {
                    $box = Find-AnswerLineBox $baseShape $lineText
                    if ($box -ne $null) { $boxes += $box }
                }
                if ($boxes.Count -eq 0) { continue }
                $desiredLeft = [double](($boxes | Measure-Object -Property Left -Minimum).Minimum)
                $desiredTop = [double](($boxes | Measure-Object -Property Top -Minimum).Minimum)
                $currentBoundLeft = [double]$shape.TextFrame.TextRange.BoundLeft
                $currentBoundTop = [double]$shape.TextFrame.TextRange.BoundTop
                $dx = $desiredLeft - $currentBoundLeft
                $dy = $desiredTop - $currentBoundTop
                if ([Math]::Abs($dx) -lt 0.5 -and [Math]::Abs($dy) -lt 0.5) { continue }
                $shape.Left = [Math]::Max(0.0, [double]$shape.Left + $dx)
                $shape.Top = [Math]::Max(0.0, [double]$shape.Top + $dy)
                $changed += 1
            } catch {
                continue
            }
        }
    }
    if ($changed -gt 0) { $pres.Save() }
    Write-Output "ZHUJI_LONG_ANSWER_ALIGNED=$changed"
}
finally {
    if ($pres -ne $null) {
        try { $pres.Close() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null } catch {}
    }
    if ($pp -ne $null) {
        try { $pp.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
'''
    env = os.environ.copy()
    env["PPTX_TO_ALIGN"] = str(pptx)
    try:
        result = subprocess.run(
            [powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            timeout=180,
        )
    except Exception as exc:
        print(f"提示：筑基长下划线答案校位失败：{exc}", file=sys.stderr)
        return
    if result.returncode != 0:
        message = result.stderr.strip() or result.stdout.strip()
        print(f"提示：筑基长下划线答案校位失败：{message}", file=sys.stderr)
        return
    output = result.stdout.strip()
    if output:
        print(f"筑基长下划线答案校位：{pptx.name}，{output}")


@isolated_generation
def generate_zuhe_jiafen_pptx(
    *,
    md_run: Path,
    template: Path,
    output_dir: Path,
    limit_groups: int | None = 1,
    group_indexes: list[int] | None = None,
    group_ranges: list[dict[str, object]] | None = None,
    progress: Callable[[str], None] | None = None,
) -> list[Path]:
    def emit(message: str) -> None:
        if progress:
            progress(message)
        else:
            print(message)

    body_md, answer_md = find_md_files(md_run)
    body_pieces = parse_body_pieces(body_md)
    answer_pieces = parse_answer_pieces(answer_md, body_pieces)
    answer_map = {wenyan.piece_key(piece): piece for piece in answer_pieces}
    tif_search_dirs = wenyan.infer_tif_search_dirs(md_run)
    source_dir = DEFAULT_SOURCE_DIR
    if source_dir.exists() and source_dir not in tif_search_dirs:
        tif_search_dirs.append(source_dir)

    output_dir.mkdir(parents=True, exist_ok=True)
    if group_ranges is not None:
        specs = []
        for index, group_range in enumerate(group_ranges, 1):
            volume = str(group_range["volume"])
            start = int(group_range["start"])
            end = int(group_range["end"])
            keys = {tuple(key) for key in group_range["unit_keys"]}  # type: ignore[arg-type]
            selected = [piece for piece in body_pieces if (piece.volume, piece.num) in keys]
            specs.append((f"{index}{volume} 第{start}-{end}练.pptx", volume, start, end, selected))
    elif group_indexes:
        specs = [GROUP_SPECS[index - 1] for index in group_indexes if 1 <= index <= len(GROUP_SPECS)]
    else:
        specs = GROUP_SPECS if limit_groups is None else GROUP_SPECS[:limit_groups]
    outputs: list[Path] = []
    emit(f"正文练习数：{len(body_pieces)}")
    emit(f"答案练习数：{len(answer_pieces)}")
    old_add_directory_content = wenyan.add_directory_content
    old_add_cover_content = wenyan.add_cover_content
    old_add_piece_title = wenyan.add_piece_title
    old_add_content = wenyan.add_content
    old_styled_body_lines = wenyan.styled_body_lines
    old_build_styled_section = wenyan.build_styled_section
    old_extract_answer_items = wenyan.extract_answer_items
    old_split_analysis = wenyan.split_analysis
    old_role_style = wenyan.role_style
    old_paragraph_properties = wenyan.paragraph_properties
    old_make_piece_slide_chunks = wenyan.make_piece_slide_chunks
    old_split_sections_with_context = wenyan.split_sections_with_context
    old_question_number_color = wenyan.QUESTION_NUMBER_COLOR
    wenyan.add_directory_content = add_zhuji_directory_content
    wenyan._zuhe_original_add_cover_content = old_add_cover_content
    wenyan.add_cover_content = add_zhuji_cover_content
    wenyan.add_piece_title = add_zhuji_piece_title
    wenyan.add_content = add_zuhe_content
    wenyan.styled_body_lines = zuhe_styled_body_lines
    wenyan.build_styled_section = zuhe_build_styled_section
    wenyan.extract_answer_items = zuhe_extract_answer_items
    wenyan.split_analysis = zuhe_split_analysis
    wenyan._zuhe_original_role_style = old_role_style
    wenyan.role_style = zuhe_role_style
    wenyan._zuhe_original_paragraph_properties = old_paragraph_properties
    wenyan.paragraph_properties = zuhe_paragraph_properties
    wenyan.make_piece_slide_chunks = zuhe_make_piece_slide_chunks
    wenyan.split_sections_with_context = zuhe_split_sections_with_context
    wenyan.QUESTION_NUMBER_COLOR = ZUHE_BLUE
    try:
        for spec in specs:
            filename, group, start, end = spec[:4]
            selected_override = spec[4] if len(spec) > 4 else None
            selected = selected_override or [piece for piece in body_pieces if start <= piece.num <= end]
            if not selected:
                emit(f"跳过，无练习：{filename}")
                continue
            output = output_dir / filename
            emit(f"生成：{output.name}，练习：{', '.join(piece.title for piece in selected)}")
            build_zhuji_deck(
                template,
                output,
                selected,
                answer_map,
                group,
                start,
                end,
                tif_search_dirs=tif_search_dirs,
                multiline_answer_min_width_pt=0.0,
            )
            outputs.append(output)
    finally:
        wenyan.add_directory_content = old_add_directory_content
        wenyan.add_cover_content = old_add_cover_content
        wenyan.add_piece_title = old_add_piece_title
        wenyan.add_content = old_add_content
        wenyan.styled_body_lines = old_styled_body_lines
        wenyan.build_styled_section = old_build_styled_section
        wenyan.extract_answer_items = old_extract_answer_items
        wenyan.split_analysis = old_split_analysis
        wenyan.role_style = old_role_style
        wenyan.paragraph_properties = old_paragraph_properties
        wenyan.make_piece_slide_chunks = old_make_piece_slide_chunks
        wenyan.split_sections_with_context = old_split_sections_with_context
        if hasattr(wenyan, "_zuhe_original_role_style"):
            delattr(wenyan, "_zuhe_original_role_style")
        if hasattr(wenyan, "_zuhe_original_add_cover_content"):
            delattr(wenyan, "_zuhe_original_add_cover_content")
        if hasattr(wenyan, "_zuhe_original_paragraph_properties"):
            delattr(wenyan, "_zuhe_original_paragraph_properties")
        wenyan.QUESTION_NUMBER_COLOR = old_question_number_color

    manifest = {
        "md_run": str(md_run),
        "template": str(template),
        "outputs": [str(path) for path in outputs],
    }
    (output_dir / "zuhe_jiafen_generation_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return outputs


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="测试：组合加分练 FBD/MD -> PPTX")
    parser.add_argument("--body", default=str(DEFAULT_BODY_FBD), help="正文 FBD 路径")
    parser.add_argument("--answer", default=str(DEFAULT_ANSWER_FBD), help="答案 FBD 路径")
    parser.add_argument("--md-run", default="", help="已有 MD run 目录；为空时先转换 FBD")
    parser.add_argument("--md-output-root", default=str(DEFAULT_MD_ROOT), help="Markdown 输出根目录")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="PPTX 模板路径")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="PPTX 输出目录")
    parser.add_argument("--limit-groups", type=int, default=1, help="生成前几个分组；0 表示全部")
    parser.add_argument("--group-index", type=int, action="append", default=[], help="只生成指定序号分组，可重复")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    md_run = Path(args.md_run) if args.md_run else convert_zuhe_fbd_to_md(
        body_fbd=Path(args.body),
        answer_fbd=Path(args.answer),
        output_root=Path(args.md_output_root),
    )
    limit_groups = None if args.limit_groups == 0 else args.limit_groups
    generate_zuhe_jiafen_pptx(
        md_run=md_run,
        template=Path(args.template),
        output_dir=Path(args.output_dir),
        limit_groups=limit_groups,
        group_indexes=args.group_index or None,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

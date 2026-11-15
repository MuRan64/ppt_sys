#!/usr/bin/env python3
"""Test generator for 语文《组合加分练》.

This first adapter intentionally reuses the mature 文言梳理 PPT engine for
slide construction, answer overlays, TIF placement, and answer animations.  The
book-specific part in this file is limited to:
- converting the two FBD files to Markdown while keeping source TIF assets;
- splitting the Markdown into practice units such as 第1练、第2练;
- distributing answer blocks to the same section titles used by the body.
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
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


DEFAULT_SOURCE_DIR = PROJECT_DIR / "高考语文组合加分练"
DEFAULT_BODY_FBD = DEFAULT_SOURCE_DIR / "组合加分练 正文.fbd"
DEFAULT_ANSWER_FBD = DEFAULT_SOURCE_DIR / "组合加分练 答案.fbd"
DEFAULT_TEMPLATE = PROJECT_DIR / "组合加分练模板.pptx"
DEFAULT_MD_ROOT = PROJECT_DIR / "out_md" / "组合加分练"
DEFAULT_OUTPUT_DIR = PROJECT_DIR / "output" / "test_zuhe_jiafen_ppt"

GROUP_SPECS = [
    ("1第一组 第1-3练.pptx", "第一组", 1, 3),
    ("2第二组 第4-6练.pptx", "第二组", 4, 6),
    ("3第三组 第7-9练.pptx", "第三组", 7, 9),
    ("4第四组 第10-12练.pptx", "第四组", 10, 12),
    ("5第五组 第13-15练.pptx", "第五组", 13, 15),
    ("6第六组 第16-18练.pptx", "第六组", 16, 18),
    ("7第七组 第19-21练.pptx", "第七组", 19, 21),
    ("8第八组 第22-24练.pptx", "第八组", 22, 24),
    ("9第九组 第25-27练.pptx", "第九组", 25, 27),
    ("10第十组 第28-30练.pptx", "第十组", 28, 30),
    ("11第十一组 第31-33练.pptx", "第十一组", 31, 33),
    ("12第十二组 第34-36练.pptx", "第十二组", 34, 36),
    ("13第十三组 第37-39练.pptx", "第十三组", 37, 39),
    ("14第十四组 第40-42练.pptx", "第十四组", 40, 42),
    ("15第十五组 第43-45练.pptx", "第十五组", 43, 45),
]

SECTION_RE = re.compile(r"^[一二三四五六七八九十]+、")
PRACTICE_TITLE_RE = re.compile(r"^第?(\d{1,3})练$")
ZUHE_TEAL = "3A7371"
ZUHE_BLUE = "00AEEF"
ZUHE_SUMMARY_ORDER = (
    "语言文字运用",
    "名篇名句默写",
    "现代文阅读Ⅰ",
    "现代文阅读Ⅱ",
    "文言文阅读",
    "古代诗歌阅读",
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
ZUHE_ANALYSIS_FRAGMENT_UNITS = 300.0
ZUHE_ANALYSIS_SINGLE_BLOCK_MAX_LINES = 13
ZUHE_MAJOR_SECTION_KEYWORDS = (
    "语言文字运用",
    "名篇名句默写",
    "现代文阅读",
    "文言文阅读",
    "古代诗歌阅读",
)
ZUHE_MATERIAL_LABEL_RE = re.compile(r"^((?:材料|文本)[一二三四五六七八九十]+[：:])\s*(.*)$")


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
    return bool(SECTION_RE.match(text) and any(keyword in text for keyword in ZUHE_MAJOR_SECTION_KEYWORDS))


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
    return bool(re.fullmatch(r"第[一二三四五六七八九十百]+组", plain(text)))


def is_practice_start(lines: list[str], index: int) -> re.Match[str] | None:
    value = plain(lines[index])
    match = re.fullmatch(r"(\d{1,3})", value)
    if not match:
        return None
    lookahead = [plain(line) for line in lines[index + 1 : index + 8] if plain(line)]
    # A practice starts with its number, then one or two page/score numbers, then
    # a first section title.  This avoids confusing exercise numbers with pieces.
    if len(lookahead) >= 3 and lookahead[0].isdigit() and lookahead[1].isdigit() and lookahead[2].startswith("一、"):
        return match
    return None


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


def remove_body_noise(lines: list[str], group: str) -> list[str]:
    cleaned: list[str] = []
    for line in normalize_section_scores(lines):
        text = plain(line)
        compact = marker(line)
        if not text and not has_fbd_tif(line):
            continue
        if text == group:
            continue
        if text in {"高考语文组合加分练", "分"}:
            continue
        if re.fullmatch(r"\d{1,3}", text):
            continue
        if compact in {"用时：分钟满分：分得分：", "第练用时：分钟满分：分得分："}:
            continue
        cleaned.append(line)
    return cleaned


def parse_body_pieces(md_path: Path) -> list[wenyan.Piece]:
    lines = [line for line in wenyan.clean_lines(md_path.read_text(encoding="utf-8")) if line.strip()]
    starts: list[tuple[int, str, int]] = []
    current_group = "第一组"
    for index, line in enumerate(lines):
        if is_group_title(line):
            current_group = plain(line)
            continue
        match = is_practice_start(lines, index)
        if match:
            starts.append((index, current_group, int(match.group(1))))

    pieces: list[wenyan.Piece] = []
    for pos, (start, group, num) in enumerate(starts):
        end = starts[pos + 1][0] if pos + 1 < len(starts) else len(lines)
        block = mark_zuhe_body_lines(remove_body_noise(lines[start + 1 : end], group))
        if not block:
            continue
        title = f"第{num}练"
        pieces.append(wenyan.Piece(volume=group, num=num, title=title, author="", lines=[title, *block]))
    return pieces


def scan_zuhe_group_units(md_run: Path) -> list[dict[str, object]]:
    """Expose the parsed MD chapter/practice sequence to the GUI grouping step."""

    body_md, _answer_md = find_md_files(md_run)
    return [
        {"volume": piece.volume, "num": piece.num, "title": piece.title, "summary": piece_summary(piece)}
        for piece in parse_body_pieces(body_md)
    ]


def chinese_group_number(value: int) -> str:
    digits = "零一二三四五六七八九"
    if value < 10:
        return digits[value]
    if value < 20:
        return "十" + (digits[value % 10] if value % 10 else "")
    if value < 100:
        return digits[value // 10] + "十" + (digits[value % 10] if value % 10 else "")
    return str(value)


def parse_answer_piece_blocks(md_path: Path) -> dict[tuple[str, int], list[str]]:
    lines = [line for line in wenyan.clean_lines(md_path.read_text(encoding="utf-8")) if line.strip()]
    current_group = "第一组"
    starts: list[tuple[int, str, int]] = []
    for index, line in enumerate(lines):
        if is_group_title(line):
            current_group = plain(line)
            continue
        match = PRACTICE_TITLE_RE.fullmatch(marker(line))
        if match:
            starts.append((index, current_group, int(match.group(1))))

    blocks: dict[tuple[str, int], list[str]] = {}
    for pos, (start, group, num) in enumerate(starts):
        end = starts[pos + 1][0] if pos + 1 < len(starts) else len(lines)
        block = [
            line
            for line in lines[start + 1 : end]
            if (plain(line) or has_fbd_tif(line)) and plain(line) != group
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


def body_section_question_nums(section_lines: list[str]) -> list[int]:
    nums: list[int] = []
    for line in section_lines:
        match = wenyan.body_item_match(line)
        if match:
            nums.append(int(match.group(1)))
    return nums


def build_answer_piece_for_body(body_piece: wenyan.Piece, answer_block: list[str] | None) -> wenyan.Piece:
    if not answer_block:
        return wenyan.Piece(body_piece.volume, body_piece.num, body_piece.title, "", [body_piece.title])
    by_num = answer_blocks_by_question(answer_block)
    lines = [body_piece.title]
    for _module, section_title, section_lines in wenyan.split_sections_with_context(body_piece):
        nums = body_section_question_nums(section_lines)
        if not nums:
            continue
        section_match = re.match(r"^([一二三四五六七八九十]+)、", plain(section_title))
        lines.append(f"{section_match.group(1)}、" if section_match else section_title)
        for num in nums:
            lines.extend(by_num.get(num, []))
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
            number = strip_zuhe_score_note(answers[answer_index])
            answer_index += 1
            line, _idx, number_filled = wenyan.fill_underlined_spans(line, [number], 0)
            filled += number_filled
            if answer_index < len(answers):
                revision = strip_zuhe_score_note(answers[answer_index])
                answer_index += 1

                def repl(match: re.Match[str]) -> str:
                    return wenyan.answer_span(revision, underline=True, pad_underline=True)

                line, count = re.subn(r"(?<=可修改为：)[_＿]{2,}", repl, line, count=1)
                filled += count
                skip_blank_only = count > 0
            result.append(line)
            continue
        result.append(line)
    return result, filled > 0


def zuhe_paragraph_properties(role: str) -> dict[str, str]:
    attrs = wenyan._zuhe_original_paragraph_properties(role)  # type: ignore[attr-defined]
    if role == "long_answer":
        attrs["indent"] = "0"
        attrs["marL"] = "0"
        attrs["algn"] = "ctr"
    if role == "material_reading":
        attrs["indent"] = "0"
        attrs["marL"] = "0"
        attrs["algn"] = "just"
    return attrs


def zuhe_role_style(role: str) -> tuple[str, float, str, bool]:
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
            styled.append(wenyan.StyledLine(line, "heading"))
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
    answer_items = zuhe_extract_answer_items(core_answer_lines)
    segments = zuhe_split_body_numbered_segments(lines)
    groups = [(num, segment_lines) for kind, num, segment_lines in segments if kind == "group" and num is not None]
    if answer_items and groups:
        styled: list[wenyan.StyledLine] = []
        matched_any = False
        preamble_filled_once = False

        for kind, num, group_lines in segments:
            if kind == "leading":
                if not preamble_filled_once:
                    preamble_lines, preamble_filled = wenyan.fill_lines_with_answer_values(
                        group_lines,
                        wenyan.extract_preamble_answers(core_answer_lines),
                    )
                    styled.extend(zuhe_styled_body_lines(preamble_lines, section_title))
                    matched_any = matched_any or preamble_filled
                    preamble_filled_once = True
                else:
                    styled.extend(zuhe_styled_body_lines(group_lines, section_title))
                continue
            if num is None:
                styled.extend(zuhe_styled_body_lines(group_lines, section_title))
                continue
            item = answer_items.get(num)
            if item is None:
                styled.extend(zuhe_styled_body_lines(group_lines, section_title))
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
            elif item.answers and is_zuhe_qa_completion_prompt(group_lines):
                filled_group, group_filled = fill_zuhe_qa_completion_answers(group_lines, item.answers)
                styled.extend(zuhe_styled_body_lines(filled_group, section_title))
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
        chunks.extend(zuhe_chunk_styled_lines(styled, max_units=13))
    return chunks


def zuhe_logical_line_blocks(lines: list[wenyan.StyledLine]) -> list[wenyan.LineBlock]:
    """Keep choices intact, but let long answer explanations paginate by paragraph."""

    blocks: list[wenyan.LineBlock] = []
    index = 0
    analysis_roles = {"analysis", "analysis_block"}
    while index < len(lines):
        line = lines[index]
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
        starts_new_question = wenyan.is_numbered_question_line(line)
        analysis_heavy_lead = current and current[-1].role in {"analysis", "analysis_block"} and used >= 4 and starts_new_question
        if analysis_heavy_lead:
            chunks.append(current)
            current = []
            used = 0

        if block.kind == "choice_body":
            next_block = blocks[index + 1] if index + 1 < len(blocks) and blocks[index + 1].kind == "choice_analysis" else None
            next_cost = wenyan.block_page_cost(next_block) if next_block else 0
            preferred_cost = cost + next_cost if next_block and cost + next_cost <= max_units else cost
            if current and (used + preferred_cost > max_units or used + cost > max_units):
                chunks.append(current)
                current = []
                used = 0
            wenyan.append_block_to_current(current, block)
            used += cost
            if next_block:
                if used + next_cost > max_units and current:
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


@isolated_generation
def generate_zuhe_jiafen_pptx(
    *,
    md_run: Path,
    template: Path,
    output_dir: Path,
    limit_groups: int | None = 1,
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
            start = int(group_range["start"])
            end = int(group_range["end"])
            group = f"第{chinese_group_number(index)}组"
            filename = f"{index}{chinese_group_number(index)}组 第{start}-{end}练.pptx"
            keys = {tuple(key) for key in group_range["unit_keys"]}  # type: ignore[arg-type]
            selected = [piece for piece in body_pieces if (piece.volume, piece.num) in keys]
            specs.append((filename, group, start, end, selected))
    else:
        base_specs = GROUP_SPECS if limit_groups is None else GROUP_SPECS[:limit_groups]
        specs = [(*spec, None) for spec in base_specs]
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
    wenyan.add_directory_content = add_zuhe_directory_content
    wenyan._zuhe_original_add_cover_content = old_add_cover_content
    wenyan.add_cover_content = add_zuhe_cover_content
    wenyan.add_piece_title = add_zuhe_piece_title
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
        for filename, group, start, end, selected_override in specs:
            selected = selected_override or [piece for piece in body_pieces if piece.volume == group and start <= piece.num <= end]
            if not selected:
                emit(f"跳过，无练习：{filename}")
                continue
            output = output_dir / filename
            emit(f"生成：{output.name}，练习：{', '.join(piece.title for piece in selected)}")
            wenyan.build_deck(
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
        "group_ranges": group_ranges,
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
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

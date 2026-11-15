#!/usr/bin/env python3
"""Experimental Markdown-to-PPTX generator for 高考语文必背古诗文.

This stays outside the GUI/main program while the new book rules are being
learned. It reuses the mature 文言梳理 generator for low-level PPTX editing,
answer embedding, TIF conversion, and answer animations, while this file owns
the book-specific parsing and template placement rules.
"""

from __future__ import annotations

import argparse
import html
import io
import json
import math
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable
from xml.etree import ElementTree as ET


PROJECT_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
TESTS_DIR = Path(__file__).resolve().parent
if str(TESTS_DIR) not in sys.path:
    sys.path.insert(0, str(TESTS_DIR))
TOOLS_DIR = PROJECT_DIR / "tools"
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

import test_md_to_ppt as wenyan  # noqa: E402
from fbd_to_md import convert_fbd_to_markdown, normalize_tif_key, read_fbd, tif_markdown_tag  # noqa: E402
from generation_state import isolated_generation  # noqa: E402


DEFAULT_MD_RUN = PROJECT_DIR / "out_md" / "必背古诗文" / "run_20260718_150544"
DEFAULT_TEMPLATE = PROJECT_DIR / "必背古诗文模板.pptx"
DEFAULT_OUTPUT_DIR = PROJECT_DIR / "output" / "test_bibei_gushiwen_ppt"
DEFAULT_BOOKLET_MD_RUN = PROJECT_DIR / "out_md" / "必背古诗文册子" / "run_20260719_134059"
DEFAULT_POCKET_MD_RUN = PROJECT_DIR / "out_md" / "必背古诗文口袋本" / "run_20260722_144611"
DEFAULT_BOOKLET_OPEN_SOURCE_DIR = PROJECT_DIR / "2026.3.3-必背古诗文"
DEFAULT_BOOKLET_OPEN_IMAGE_MD_RUN = PROJECT_DIR / "out_md" / "必背古诗文册子开放性图片" / "run_latest"

BOOK_TITLE = "高考语文必背古诗文"
BOOK_LABEL = "新课标必背60篇"
FIRST_GROUP_NAME = "1 篇目1-6+「6+1训练赢」.pptx"
BIBEI_MAIN_GROUP_SPECS = [(index, start, start + 5) for index, start in enumerate(range(1, 61, 6), 1)]

VOLUME_ALIASES = {
    "必修上": "必修上册",
    "必修下": "必修下册",
    "选择性必修上": "选择性必修上册",
    "选择性必修中": "选择性必修中册",
    "选择性必修下": "选择性必修下册",
}
VOLUME_NAMES = set(VOLUME_ALIASES.values())

GREEN = "679163"
BLUE = "00AEEF"
BLACK = "000000"
WHITE = "FFFFFF"

WENYAN_ROLE_STYLE = wenyan.role_style

BIBEI_SUBHEADING_LINES = {
    "真题再现",
    "经典常考",
    "第一组",
    "第二组",
    "第三组",
    "第四组",
    "第五组",
}
BIBEI_FORCE_NEW_CHUNK_SUBHEADINGS = {"经典常考", "第一组", "第二组", "第三组", "第四组", "第五组"}
BIBEI_TRAINING_MARKER = "6+1训练赢"
BIBEI_TRAINING_GROUP_LINES = {"第一组", "第二组", "第三组", "第四组", "第五组"}
BIBEI_ANY_GROUP_RE = re.compile(r"^(?:第[一二三四五六七八九十百]+组|第\d+组)$")
BIBEI_RUSH_DAY_RE = re.compile(r"^(?:第[一二三四五六七八九十百]+天|第\d+天)$")
BIBEI_DEFAULT_PAGE_UNITS = 10
BIBEI_MOXIE_PAGE_UNITS = 11
BIBEI_EXPLANATION_PAGE_UNITS = 10
BIBEI_RUSH_PAGE_UNITS = 11
BIBEI_TIF_MAX_WIDTH_PX = 1600
BIBEI_TIF_MAX_W_IN = 12.05
BIBEI_TIF_MAX_H_IN = 5.10
BIBEI_TIF_FOOTER_TOP_IN = 6.36
BIBEI_TIF_TITLE_SHORT_Y = 1.48
BIBEI_TIF_TITLE_MID_Y = 1.34
BIBEI_TIF_TITLE_TALL_Y = 1.27
BIBEI_CONTENT_BOX = (0.70, 0.55, 11.92, 5.78)
BOOKLET_CONTENT_BOX = (0.70, 0.48, 11.92, 5.86)
OPEN_TOPIC_PAGE_UNITS = 11
OPEN_TOPIC_IMAGE_GAP_IN = 0.18
OPEN_TOPIC_IMAGE_MAX_H_IN = 1.28
OPEN_TOPIC_IMAGE_LABEL_H_IN = 0.26
OPEN_TOPIC_IMAGE_TOP_PAD_IN = 0.14
OPEN_TOPIC_IMAGE_BOTTOM_PAD_IN = 0.16
OPEN_TOPIC_INLINE_ANSWER_Y_OFFSET_IN = 0.035

BIBEI_TIF_TEXT_ALIASES = {
    "rou-hb.tif": "輮",
    "rou-hb-cs.tif": "輮",
    "chuang-hb.tif": "摐",
    "chuang-hb-cs.tif": "摐",
    "ji-hb.tif": "羁",
    "ji-hb-cs.tif": "羁",
    "xiang-hb.tif": "襄",
    "xiang-hb-cs.tif": "襄",
    "yan3-hb.tif": "巘",
    "yan3-hb-cs.tif": "巘",
    "jue1-k.tif": "駃",
    "jue1-hb.tif": "駃",
    "ti1-k.tif": "騠",
    "ti1-hb.tif": "騠",
    "kuang-hb.tif": "纩",
    "kuang-hb-cs.tif": "纩",
    "fei-hb.tif": "騑",
    "fei-hb-cs.tif": "騑",
}

OPEN_TOPIC_NUMERAL_MAP = {
    "一": 1,
    "二": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
}
OPEN_TOPIC_MAX_NUM = 8


def bibei_paragraph_properties(role: str) -> dict[str, str]:
    attrs = {
        "marL": "0",
        "indent": "0",
        "algn": "l",
        "eaLnBrk": "1",
        "latinLnBrk": "0",
        "hangingPunct": "0",
    }
    if role in {"reading", "extended_reading"}:
        attrs["algn"] = "just"
    if role == "source_note":
        attrs["algn"] = "r"
    if role == "booklet_moxie_body":
        attrs["indent"] = str(wenyan.inch(0.44))
    if role == "booklet_moxie_body_cont":
        attrs["indent"] = "0"
    if role in {"center_heading", "rush_day", "rush_group", "booklet_moxie_title"}:
        attrs["algn"] = "ctr"
    if role == "open_image_labels":
        attrs["algn"] = "ctr"
    return attrs


def bibei_role_style(role: str) -> tuple[str, float, str, bool]:
    if role == "rush_day":
        return "黑体", 24, BLACK, False
    if role == "rush_group":
        return "黑体", 24, BLUE, False
    if role == "booklet_moxie_title":
        return "黑体", 24, BLUE, False
    if role == "booklet_moxie_subtitle":
        return "黑体", 24, BLACK, False
    if role == "booklet_moxie_body":
        return WENYAN_ROLE_STYLE("body")
    if role == "booklet_moxie_body_cont":
        return WENYAN_ROLE_STYLE("body")
    if role == "open_image_labels":
        return "宋体", 18, BLACK, False
    return WENYAN_ROLE_STYLE(role)


@dataclass
class ParseStats:
    body_count: int
    answer_count: int
    first_titles: list[str]


def plain(text: str) -> str:
    value = wenyan.strip_inline_tags(html.unescape(text)).strip()
    value = re.sub(r"_+", "", value).strip()
    return value


def title_plain(text: str) -> str:
    value = wenyan.strip_inline_tags(html.unescape(text)).strip()
    value = re.sub(r"_+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def nonempty_after(lines: list[str], index: int, count: int = 4) -> list[str]:
    result: list[str] = []
    for line in lines[index + 1 :]:
        if line.strip():
            result.append(line)
        if len(result) >= count:
            break
    return result


def is_body_piece_start(lines: list[str], index: int) -> re.Match[str] | None:
    line = plain(lines[index])
    if not re.fullmatch(r"\d{1,3}", line):
        return None
    lookahead = [plain(item) for item in nonempty_after(lines, index, 5)]
    if len(lookahead) < 2:
        return None
    title = lookahead[0]
    exam_or_page = lookahead[1]
    if title in {BOOK_TITLE, BOOK_LABEL}:
        return None
    if re.fullmatch(r"\d+年\d+考", exam_or_page):
        return re.match(r"^(\d{1,3})$", line)
    if exam_or_page.isdigit() and len(lookahead) > 2 and lookahead[2].startswith("一、"):
        return re.match(r"^(\d{1,3})$", line)
    return None


def extract_volume_from_block(block: list[str], fallback: str = "必修上册") -> str:
    for raw in block:
        value = title_plain(raw)
        match = re.search(r"（([^）]+)）", value)
        if not match:
            continue
        label = match.group(1).strip()
        if label in VOLUME_NAMES:
            return label
        if label in VOLUME_ALIASES:
            return VOLUME_ALIASES[label]
    return fallback


def is_volume_title_footer(line: str) -> bool:
    value = title_plain(line)
    return bool(re.match(r"^（[^）]+）\s*.+", value))


def is_global_noise(line: str) -> bool:
    if "<img" in line:
        return False
    value = title_plain(line)
    if not value:
        return True
    if value in {BOOK_TITLE, BOOK_LABEL, "答案链接〉〉〉P", "答 案 链 接〉〉〉P", "6+1训练赢"}:
        return True
    if re.fullmatch(r"[\d\s*×xX.]+", value):
        return True
    if value.startswith("时间：") or value.startswith("得分："):
        return True
    return False


def normalize_body_numbering(line: str) -> str:
    if re.match(r"^\s*\d{1,2}[\.、]", wenyan.strip_inline_tags(line)):
        return line
    return re.sub(r"^(\s*)(\d{1,2})(?=\D)", r"\1\2.", line, count=1)


def parse_body_pieces(md_path: Path) -> list[wenyan.Piece]:
    lines = wenyan.clean_lines(md_path.read_text(encoding="utf-8"))
    starts: list[tuple[int, int]] = []
    for index in range(len(lines)):
        match = is_body_piece_start(lines, index)
        if match:
            starts.append((index, int(match.group(1))))

    pieces: list[wenyan.Piece] = []
    previous_volume = "必修上册"
    for pos, (start, num) in enumerate(starts):
        end = starts[pos + 1][0] if pos + 1 < len(starts) else len(lines)
        block = [line for line in lines[start:end] if line.strip()]
        if len(block) < 3:
            continue
        title = title_plain(block[1])
        second = plain(block[2])
        if re.fullmatch(r"\d+年\d+考", second):
            exam_count = second
            content_start = 3
        else:
            exam_count = ""
            content_start = 2
        volume = extract_volume_from_block(block, previous_volume)
        previous_volume = volume

        if len(block) > content_start and plain(block[content_start]).isdigit():
            content_start += 1
        content_lines: list[str] = []
        for raw in block[content_start:]:
            if title_plain(raw) in {"附赠选修12篇", "150组冲刺练"}:
                break
            if is_global_noise(raw) or is_volume_title_footer(raw):
                continue
            content_lines.append(normalize_body_numbering(raw))
        piece_lines = [title, *([exam_count] if exam_count else []), *content_lines]
        pieces.append(wenyan.Piece(volume=volume, num=num, title=title, author=exam_count, lines=piece_lines))
    return pieces


def normalize_answer_volume(value: str, current: str) -> str:
    value = title_plain(value)
    if value in VOLUME_ALIASES:
        return VOLUME_ALIASES[value]
    if value in VOLUME_NAMES:
        return value
    return current


def is_answer_piece_start(lines: list[str], index: int) -> re.Match[str] | None:
    value = title_plain(lines[index])
    match = re.match(r"^(\d{1,3})\.\s*(.+)$", value)
    if not match:
        return None
    lookahead = [title_plain(item) for item in nonempty_after(lines, index, 2)]
    if lookahead and lookahead[0].startswith("一、"):
        return match
    return None


def parse_answer_pieces(md_path: Path) -> list[wenyan.Piece]:
    lines = wenyan.clean_lines(md_path.read_text(encoding="utf-8"))
    current_volume = "必修上册"
    starts: list[tuple[int, str, int, str]] = []
    for index, line in enumerate(lines):
        value = title_plain(line)
        current_volume = normalize_answer_volume(value, current_volume)
        match = is_answer_piece_start(lines, index)
        if match:
            starts.append((index, current_volume, int(match.group(1)), match.group(2).strip()))

    pieces: list[wenyan.Piece] = []
    for pos, (start, volume, num, title) in enumerate(starts):
        end = starts[pos + 1][0] if pos + 1 < len(starts) else len(lines)
        block = [line for line in lines[start + 1 : end] if line.strip()]
        block = [line for line in block if title_plain(line) not in {BOOK_TITLE, BOOK_LABEL, "答案解析"}]
        block = [normalize_answer_line(line) for line in block]
        pieces.append(wenyan.Piece(volume=volume, num=num, title=title_plain(title), author="", lines=[title_plain(title), *block]))
    return pieces


def canonical_bibei_piece_title(title: str) -> str:
    value = title_plain(title)
    value = value.replace("（", "(").replace("）", ")")
    value = re.sub(r"[\s　]+", "", value)
    return value


def build_bibei_answer_piece_map(
    body_pieces: list[wenyan.Piece],
    answer_pieces: list[wenyan.Piece],
) -> dict[tuple[str, int], wenyan.Piece]:
    answer_map = {wenyan.piece_key(piece): piece for piece in answer_pieces}
    by_num_title: dict[tuple[int, str], wenyan.Piece] = {}
    by_num: dict[int, list[wenyan.Piece]] = {}
    for piece in answer_pieces:
        by_num_title[(piece.num, canonical_bibei_piece_title(piece.title))] = piece
        by_num.setdefault(piece.num, []).append(piece)

    for body_piece in body_pieces:
        body_key = wenyan.piece_key(body_piece)
        if body_key in answer_map:
            continue
        title_key = (body_piece.num, canonical_bibei_piece_title(body_piece.title))
        candidate = by_num_title.get(title_key)
        if candidate is None:
            same_num = by_num.get(body_piece.num, [])
            if len(same_num) == 1:
                candidate = same_num[0]
        if candidate is not None:
            answer_map[body_key] = candidate
    return answer_map


def normalize_answer_line(line: str) -> str:
    line = re.sub(r"^\s*[_＿]{2,}\s*(?=\d{1,3}[\.、])", "", line)
    return line


def restyle_bibei_styled_lines(lines: list[wenyan.StyledLine]) -> list[wenyan.StyledLine]:
    styled: list[wenyan.StyledLine] = []
    for line in lines:
        plain_text = title_plain(line.text)
        if plain_text in BIBEI_SUBHEADING_LINES or BIBEI_ANY_GROUP_RE.fullmatch(plain_text):
            meta = dict(line.meta or {})
            if plain_text in BIBEI_FORCE_NEW_CHUNK_SUBHEADINGS:
                meta["force_new_chunk"] = True
            styled.append(wenyan.StyledLine(line.text, "subheading", meta))
            continue
        styled.append(line)
    return styled


def mark_rush_question_number_blue(text: str) -> str:
    if text.lstrip().startswith('<span class="blue-text">'):
        return text
    return re.sub(
        r"^(\s*)(\d+[\.、])",
        lambda match: f'{match.group(1)}<span class="blue-text">{match.group(2)}</span>',
        text,
        count=1,
    )


def restyle_bibei_rush_styled_lines(lines: list[wenyan.StyledLine]) -> list[wenyan.StyledLine]:
    styled: list[wenyan.StyledLine] = []
    for line in lines:
        plain_text = title_plain(line.text)
        if BIBEI_RUSH_DAY_RE.fullmatch(plain_text):
            styled.append(wenyan.StyledLine(plain_text, "rush_day", dict(line.meta or {})))
            continue
        if BIBEI_ANY_GROUP_RE.fullmatch(plain_text):
            meta = dict(line.meta or {})
            styled.append(wenyan.StyledLine(plain_text, "rush_group", meta))
            continue
        if line.role in {"body", "reading", "extended_reading"} and re.match(r"^\s*\d+[\.、]", wenyan.strip_inline_tags(line.text).strip()):
            styled.append(wenyan.StyledLine(mark_rush_question_number_blue(line.text), line.role, dict(line.meta or {})))
            continue
        styled.append(line)
    return styled


def bibei_section_page_units(title: str) -> int:
    plain_title = title_plain(title)
    if "开放性默写专题练" in plain_title:
        return 5
    if "解释并掌握" in plain_title or "重难点字词" in plain_title:
        return BIBEI_EXPLANATION_PAGE_UNITS
    if "150组冲刺练" in plain_title:
        return BIBEI_RUSH_PAGE_UNITS
    if "理解性默写" in plain_title:
        return BIBEI_MOXIE_PAGE_UNITS
    return BIBEI_DEFAULT_PAGE_UNITS


def split_bibei_training_answer_lines(answer_lines: list[str] | None) -> tuple[list[str] | None, list[str] | None]:
    if not answer_lines:
        return answer_lines, None
    marker_index = next((i for i, line in enumerate(answer_lines) if title_plain(line) == BIBEI_TRAINING_MARKER), None)
    if marker_index is None:
        return answer_lines, None
    normal_lines = [line for line in answer_lines[:marker_index] if title_plain(line) != BIBEI_TRAINING_MARKER]
    training_lines = [line for line in answer_lines[marker_index + 1 :] if title_plain(line) != BIBEI_TRAINING_MARKER]
    return normal_lines, training_lines


def bibei_answer_section_map(answer_piece: wenyan.Piece | None) -> dict[str, list[str]]:
    section_map = wenyan.answer_section_map(answer_piece)
    moxie_key = wenyan.section_context_key("", "二、理解性默写。")
    training_key = wenyan.section_context_key("", "三、易错字词强化练习。")
    normal_lines, training_lines = split_bibei_training_answer_lines(section_map.get(moxie_key))
    if training_lines:
        section_map[moxie_key] = normal_lines or []
        section_map[training_key] = training_lines
    return section_map


def split_bibei_training_groups(lines: list[str]) -> tuple[list[str], list[tuple[str, list[str]]]]:
    leading: list[str] = []
    groups: list[tuple[str, list[str]]] = []
    current_title: str | None = None
    current_lines: list[str] = []

    for line in lines:
        value = title_plain(line)
        if value in BIBEI_TRAINING_GROUP_LINES or BIBEI_ANY_GROUP_RE.fullmatch(value):
            if current_title is None:
                leading = current_lines
            else:
                groups.append((current_title, current_lines))
            current_title = value
            current_lines = [line]
            continue
        current_lines.append(line)

    if current_title is None:
        leading = current_lines
    else:
        groups.append((current_title, current_lines))
    return leading, groups


def build_bibei_training_group_lines(
    group_lines: list[str],
    answer_lines: list[str] | None,
    section_title: str,
) -> list[wenyan.StyledLine]:
    answer_items = wenyan.extract_answer_items(answer_lines)
    leading, groups = wenyan.split_body_numbered_groups(group_lines)
    styled = wenyan.styled_body_lines(leading, section_title)

    for num, numbered_lines in groups:
        item = answer_items.get(num)
        if item is None:
            styled.extend(wenyan.styled_body_lines(numbered_lines, section_title))
            continue

        if wenyan.section_has_fillable_blanks(numbered_lines):
            filled_lines, _filled = wenyan.fill_lines_with_answer_values(numbered_lines, item.answers)
            styled.extend(wenyan.styled_body_lines(filled_lines, section_title))
        else:
            styled.extend(wenyan.styled_body_lines(numbered_lines, section_title))
            for answer in item.answers:
                styled.append(wenyan.StyledLine(answer, "answer"))

        for analysis in item.analyses:
            styled.append(wenyan.StyledLine(analysis, "analysis"))
    return styled


def build_bibei_training_section(lines: list[str], answer_lines: list[str] | None) -> list[wenyan.StyledLine]:
    if not answer_lines:
        return restyle_bibei_styled_lines(wenyan.styled_body_lines(lines, lines[0] if lines else ""))

    section_title = lines[0] if lines else ""
    leading, body_groups = split_bibei_training_groups(lines)
    _answer_leading, answer_groups = split_bibei_training_groups(answer_lines)
    answer_by_group = {title: group_lines for title, group_lines in answer_groups}

    styled: list[wenyan.StyledLine] = []
    styled.extend(wenyan.styled_body_lines(leading, section_title))
    for group_title, group_lines in body_groups:
        styled.extend(build_bibei_training_group_lines(group_lines, answer_by_group.get(group_title), section_title))
    return restyle_bibei_styled_lines(styled)


def build_bibei_rush_section(lines: list[str], answer_lines: list[str] | None) -> list[wenyan.StyledLine]:
    return restyle_bibei_rush_styled_lines(build_bibei_training_section(lines, answer_lines))


def chunk_bibei_rush_lines(lines: list[wenyan.StyledLine], max_units: int) -> list[list[wenyan.StyledLine]]:
    chunks: list[list[wenyan.StyledLine]] = []
    leading: list[wenyan.StyledLine] = []
    pending_day: list[wenyan.StyledLine] = []
    current_group: list[wenyan.StyledLine] = []
    saw_group = False

    def flush_group() -> None:
        nonlocal current_group
        if not current_group:
            return
        chunks.extend(wenyan.chunk_styled_lines(current_group, max_units=max_units))
        current_group = []

    for line in lines:
        if line.role == "rush_day":
            flush_group()
            pending_day = [line]
            saw_group = False
            continue
        if line.role == "rush_group":
            flush_group()
            current_group = [*leading, *pending_day, line]
            leading = []
            pending_day = []
            saw_group = True
            continue
        if saw_group or current_group:
            current_group.append(line)
        else:
            leading.append(line)

    flush_group()
    if leading or pending_day:
        chunks.extend(wenyan.chunk_styled_lines([*leading, *pending_day], max_units=max_units))
    return chunks


def is_open_topic_section(body_piece: wenyan.Piece, title: str) -> bool:
    return (
        body_piece.volume == "册子"
        and title_plain(body_piece.title).startswith("专题")
        and "开放性默写专题练" in title_plain(title)
    )


def open_figure_labels(line: str) -> list[str]:
    return re.findall(r"图\d+", plain(line))


def restyle_open_topic_styled_lines(lines: list[wenyan.StyledLine]) -> list[wenyan.StyledLine]:
    converted: list[wenyan.StyledLine] = []
    for line in lines:
        if line_has_open_topic_image(line.text):
            meta = dict(line.meta or {})
            meta["tif_names"] = extract_tif_names_from_text(line.text)
            converted.append(wenyan.StyledLine(line.text, "open_image_row", meta))
            continue
        if is_open_figure_placeholder_line(line.text):
            converted.append(wenyan.StyledLine(" ".join(open_figure_labels(line.text)), "open_image_labels", dict(line.meta or {})))
            continue
        converted.append(line)

    reordered: list[wenyan.StyledLine] = []
    index = 0
    while index < len(converted):
        line = converted[index]
        if line.role != "open_image_row":
            reordered.append(line)
            index += 1
            continue
        image_group = [line]
        index += 1
        if index < len(converted) and converted[index].role == "open_image_labels":
            image_group.append(converted[index])
            index += 1
        analyses: list[wenyan.StyledLine] = []
        while index < len(converted) and (converted[index].role or "body") == "analysis":
            analyses.append(converted[index])
            index += 1
        reordered.extend(analyses)
        reordered.extend(image_group)
    return reordered


def is_open_numbered_question_line(line: wenyan.StyledLine) -> bool:
    if (line.role or "body") not in {"body", "reading", "extended_reading"}:
        return False
    return bool(re.match(r"^\s*\d+[\.、]\s*", wenyan.strip_inline_tags(line.text).strip()))


def open_topic_line_page_cost(line: wenyan.StyledLine) -> int:
    role = line.role or "body"
    if role == "open_image_row":
        tif_count = len(line.meta.get("tif_names", [])) if line.meta else len(extract_tif_names_from_text(line.text))
        return 5 if tif_count >= 3 else 4
    if role == "open_image_labels":
        return 1
    return max(1, bibei_visual_line_count(line, BOOKLET_CONTENT_BOX[2]))


def chunk_open_topic_lines(lines: list[wenyan.StyledLine], max_units: int = OPEN_TOPIC_PAGE_UNITS) -> list[list[wenyan.StyledLine]]:
    leading: list[wenyan.StyledLine] = []
    groups: list[list[wenyan.StyledLine]] = []
    current: list[wenyan.StyledLine] = []
    for line in lines:
        if is_open_numbered_question_line(line):
            if current:
                groups.append(current)
            current = [line]
            continue
        if current:
            current.append(line)
        else:
            leading.append(line)
    if current:
        groups.append(current)

    chunks: list[list[wenyan.StyledLine]] = []
    current_chunk: list[wenyan.StyledLine] = list(leading)
    used = sum(open_topic_line_page_cost(line) for line in current_chunk)
    for group in groups:
        group_cost = sum(open_topic_line_page_cost(line) for line in group)
        if current_chunk and used + group_cost > max_units:
            chunks.append(current_chunk)
            current_chunk = []
            used = 0
        current_chunk.extend(group)
        used += group_cost
    if current_chunk:
        chunks.append(current_chunk)
    return chunks


def make_bibei_piece_slide_chunks(body_piece: wenyan.Piece, answer_piece: wenyan.Piece | None) -> list[list[wenyan.StyledLine]]:
    answer_map = bibei_answer_section_map(answer_piece)
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
            if module_answer_lines and any("【解析】" in line for line in module_answer_lines):
                answer_lines = module_answer_lines
        is_rush_section = "150组冲刺练" in title_plain(title)
        if is_rush_section:
            section_lines = build_bibei_rush_section(lines, answer_lines)
        elif title_plain(title).startswith("三、易错字词强化练习"):
            section_lines = build_bibei_training_section(lines, answer_lines)
        else:
            section_lines, _matched_answers = wenyan.build_styled_section(lines, answer_lines)
            section_lines = restyle_bibei_styled_lines(section_lines)
        is_open_section = is_open_topic_section(body_piece, title)
        if is_open_section:
            section_lines = restyle_open_topic_styled_lines(section_lines)
        is_booklet_moxie_section = body_piece.volume == "册子" and title_plain(title).startswith("一、默写挖空练")
        if is_booklet_moxie_section:
            section_lines = [
                wenyan.StyledLine(line.text, "booklet_moxie_title", line.meta)
                if index == 0 and title_plain(line.text).startswith("一、默写挖空练")
                else line
                for index, line in enumerate(section_lines)
            ]
        if body_piece.volume == "册子":
            section_lines = split_booklet_moxie_styled_lines(section_lines, moxie_roles=is_booklet_moxie_section)
        if pending_module_header:
            section_lines = [wenyan.StyledLine(pending_module_header, "module_header")] + section_lines
            pending_module_header = None
        if is_rush_section:
            chunks.extend(chunk_bibei_rush_lines(section_lines, max_units=bibei_section_page_units(title)))
        elif is_booklet_moxie_section:
            chunks.extend(chunk_booklet_moxie_lines(section_lines))
        elif is_open_section:
            chunks.extend(chunk_open_topic_lines(section_lines))
        else:
            chunks.extend(wenyan.chunk_styled_lines(section_lines, max_units=bibei_section_page_units(title)))
    return chunks


def add_textbox_runs(
    root: ET.Element,
    *,
    x: float,
    y: float,
    w: float,
    h: float,
    runs: list[tuple[str, str, float, str, bool]],
    name: str,
    align: str = "l",
) -> int:
    shape_id = wenyan.next_shape_id(root)
    sp = ET.Element(wenyan.qn(wenyan.P_NS, "sp"))
    nv_sp_pr = ET.SubElement(sp, wenyan.qn(wenyan.P_NS, "nvSpPr"))
    ET.SubElement(nv_sp_pr, wenyan.qn(wenyan.P_NS, "cNvPr"), {"id": str(shape_id), "name": f"{name} {shape_id}"})
    ET.SubElement(nv_sp_pr, wenyan.qn(wenyan.P_NS, "cNvSpPr"), {"txBox": "1"})
    ET.SubElement(nv_sp_pr, wenyan.qn(wenyan.P_NS, "nvPr"))

    sp_pr = ET.SubElement(sp, wenyan.qn(wenyan.P_NS, "spPr"))
    xfrm = ET.SubElement(sp_pr, wenyan.qn(wenyan.A_NS, "xfrm"))
    ET.SubElement(xfrm, wenyan.qn(wenyan.A_NS, "off"), {"x": str(wenyan.inch(x)), "y": str(wenyan.inch(y))})
    ET.SubElement(xfrm, wenyan.qn(wenyan.A_NS, "ext"), {"cx": str(wenyan.inch(w)), "cy": str(wenyan.inch(h))})
    geom = ET.SubElement(sp_pr, wenyan.qn(wenyan.A_NS, "prstGeom"), {"prst": "rect"})
    ET.SubElement(geom, wenyan.qn(wenyan.A_NS, "avLst"))
    ET.SubElement(sp_pr, wenyan.qn(wenyan.A_NS, "noFill"))
    ln = ET.SubElement(sp_pr, wenyan.qn(wenyan.A_NS, "ln"))
    ET.SubElement(ln, wenyan.qn(wenyan.A_NS, "noFill"))

    tx_body = ET.SubElement(sp, wenyan.qn(wenyan.P_NS, "txBody"))
    ET.SubElement(tx_body, wenyan.qn(wenyan.A_NS, "bodyPr"), {"wrap": "square", "lIns": "0", "tIns": "0", "rIns": "0", "bIns": "0"})
    ET.SubElement(tx_body, wenyan.qn(wenyan.A_NS, "lstStyle"))
    paragraph = ET.SubElement(tx_body, wenyan.qn(wenyan.A_NS, "p"))
    ET.SubElement(paragraph, wenyan.qn(wenyan.A_NS, "pPr"), {"algn": align})
    for text, font, size, color, bold in runs:
        wenyan.append_run(paragraph, text, font=font, size=size, color=color, bold=bold)
    size = runs[-1][2] if runs else 24
    ET.SubElement(paragraph, wenyan.qn(wenyan.A_NS, "endParaRPr"), {"lang": "zh-CN", "sz": str(int(size * 100))})
    wenyan.append_to_sp_tree(root, sp)
    return shape_id


def add_bibei_cover_content(
    root: ET.Element,
    volume: str,
    start: int,
    end: int,
    *,
    subtitle: str | None = None,
    include_training_label: bool = True,
) -> None:
    add_textbox_runs(
        root,
        x=3.10,
        y=3.10,
        w=9.55,
        h=1.02,
        runs=[(BOOK_TITLE, "Songti SC Black", 70, BLACK, True)],
        name="封面标题",
        align="ctr",
    )
    display_subtitle = subtitle or f"篇目{start}-{end}" + (" +「6+1训练赢」" if include_training_label else "")
    add_textbox_runs(
        root,
        x=5.82,
        y=4.78,
        w=5.45,
        h=0.48,
        runs=[(display_subtitle, "黑体", 26, BLACK, True)],
        name="封面副标题",
        align="ctr",
    )


def add_bibei_directory_item(
    root: ET.Element,
    rels_root: ET.Element,
    *,
    piece: wenyan.Piece,
    slide_number: int,
    x: float,
    y: float,
    w: float,
) -> None:
    shape_id = wenyan.next_shape_id(root)
    sp = ET.Element(wenyan.qn(wenyan.P_NS, "sp"))
    nv_sp_pr = ET.SubElement(sp, wenyan.qn(wenyan.P_NS, "nvSpPr"))
    c_nv_pr = ET.SubElement(nv_sp_pr, wenyan.qn(wenyan.P_NS, "cNvPr"), {"id": str(shape_id), "name": f"目录项 {piece.num}"})
    rel_id = wenyan.add_slide_jump_relationship(rels_root, slide_number)
    ET.SubElement(c_nv_pr, wenyan.qn(wenyan.A_NS, "hlinkClick"), {wenyan.qn(wenyan.R_NS, "id"): rel_id, "action": "ppaction://hlinksldjump"})
    ET.SubElement(nv_sp_pr, wenyan.qn(wenyan.P_NS, "cNvSpPr"), {"txBox": "1"})
    ET.SubElement(nv_sp_pr, wenyan.qn(wenyan.P_NS, "nvPr"))

    sp_pr = ET.SubElement(sp, wenyan.qn(wenyan.P_NS, "spPr"))
    xfrm = ET.SubElement(sp_pr, wenyan.qn(wenyan.A_NS, "xfrm"))
    ET.SubElement(xfrm, wenyan.qn(wenyan.A_NS, "off"), {"x": str(wenyan.inch(x)), "y": str(wenyan.inch(y))})
    ET.SubElement(xfrm, wenyan.qn(wenyan.A_NS, "ext"), {"cx": str(wenyan.inch(w)), "cy": str(wenyan.inch(0.55))})
    geom = ET.SubElement(sp_pr, wenyan.qn(wenyan.A_NS, "prstGeom"), {"prst": "rect"})
    ET.SubElement(geom, wenyan.qn(wenyan.A_NS, "avLst"))
    ET.SubElement(sp_pr, wenyan.qn(wenyan.A_NS, "noFill"))
    ln = ET.SubElement(sp_pr, wenyan.qn(wenyan.A_NS, "ln"))
    ET.SubElement(ln, wenyan.qn(wenyan.A_NS, "noFill"))

    tx_body = ET.SubElement(sp, wenyan.qn(wenyan.P_NS, "txBody"))
    ET.SubElement(tx_body, wenyan.qn(wenyan.A_NS, "bodyPr"), {"wrap": "square", "lIns": "0", "tIns": "0", "rIns": "0", "bIns": "0"})
    ET.SubElement(tx_body, wenyan.qn(wenyan.A_NS, "lstStyle"))
    paragraph = ET.SubElement(tx_body, wenyan.qn(wenyan.A_NS, "p"))
    ET.SubElement(paragraph, wenyan.qn(wenyan.A_NS, "pPr"), {"algn": "l"})
    display_title = piece.title.replace("分册强化练", "") if piece.volume == "册子" else piece.title
    prefix = "" if piece.volume == "册子" and re.match(r"^专题\d+", display_title) else f"{piece.num} "
    if prefix:
        wenyan.append_run(paragraph, prefix, font="黑体", size=32, color=BLACK, bold=True)
    wenyan.append_run(paragraph, display_title, font="黑体", size=32, color=BLACK, bold=True)
    ET.SubElement(paragraph, wenyan.qn(wenyan.A_NS, "endParaRPr"), {"lang": "zh-CN", "sz": "3200"})
    wenyan.append_to_sp_tree(root, sp)


def add_bibei_directory_content(
    root: ET.Element,
    rels_root: ET.Element,
    pieces: list[wenyan.Piece],
    title_slide_numbers: dict[tuple[str, int], int],
    *,
    include_training_label: bool = True,
) -> None:
    for sp in wenyan.text_shapes(root):
        text = wenyan.get_shape_text(sp).strip()
        if text and "各篇目点击即可直达" not in text:
            wenyan.set_shape_text(sp, "")

    if len(pieces) > 6 and not include_training_label:
        positions = [
            (2.60, 0.98, 4.35),
            (7.58, 0.98, 4.35),
            (2.60, 2.12, 4.35),
            (7.58, 2.12, 4.35),
            (2.60, 3.26, 4.35),
            (7.58, 3.26, 4.35),
            (2.60, 4.40, 4.35),
            (7.58, 4.40, 4.35),
        ]
    else:
        positions = [
            (3.05, 1.08, 4.10),
            (3.05, 2.43, 4.30),
            (3.05, 3.76, 4.55),
            (8.15, 1.08, 3.75),
            (8.15, 2.43, 4.25),
            (8.15, 3.76, 4.35),
        ]
    for piece, (x, y, w) in zip(pieces, positions):
        add_bibei_directory_item(root, rels_root, piece=piece, slide_number=title_slide_numbers[wenyan.piece_key(piece)], x=x, y=y, w=w)
        wenyan.append_dashed_line(root, x=x - 0.32, y=y + 0.47, w=w + 0.25, name=f"目录虚线 {piece.num}")
    if include_training_label:
        add_textbox_runs(
            root,
            x=5.30,
            y=5.35,
            w=4.15,
            h=0.55,
            runs=[("6+1训练赢", "黑体", 26, BLACK, True)],
            name="目录6+1",
            align="ctr",
        )
        wenyan.append_dashed_line(root, x=5.15, y=5.78, w=4.45, name="目录6+1虚线")


def bibei_piece_title_font_size(title: str) -> float:
    units = wenyan.text_units(title)
    if units <= 10.5:
        return 80
    if units <= 13:
        return 72
    if units <= 16:
        return 64
    return 54


def add_bibei_piece_title(root: ET.Element, piece: wenyan.Piece) -> None:
    wenyan.clear_text_shapes(root)
    if piece.volume == "册子":
        display_title = piece.title.replace("分册强化练", "")
        add_textbox_runs(
            root,
            x=1.35,
            y=1.54,
            w=10.70,
            h=0.95,
            runs=[(display_title, "黑体", 46, BLACK, True)],
            name="小册子标题",
            align="ctr",
        )
        if piece.author:
            add_textbox_runs(
                root,
                x=0.82,
                y=2.78,
                w=11.25,
                h=1.65,
                runs=[(piece.author, "宋体", 18, BLACK, False)],
                name="小册子内容",
                align="l",
            )
        add_textbox_runs(
            root,
            x=0.40,
            y=6.43,
            w=2.70,
            h=0.40,
            runs=[("分 册 强 化 练", "微软雅黑", 20, WHITE, False)],
            name="小册子标题册别",
            align="l",
        )
        return
    display_title = piece.title.replace(" ", "　") if piece.title == "登 高" else piece.title
    title_size = bibei_piece_title_font_size(display_title)
    if piece.title in {"150组冲刺练"}:
        add_textbox_runs(
            root,
            x=1.15,
            y=2.28,
            w=10.30,
            h=1.45,
            runs=[(display_title, "方正小标宋_GBK", title_size, BLACK, True)],
            name="篇目标题",
            align="ctr",
        )
        add_textbox_runs(
            root,
            x=0.40,
            y=6.43,
            w=2.30,
            h=0.40,
            runs=[(piece.volume, "微软雅黑", 20, WHITE, False)],
            name="篇目标题册别",
            align="l",
        )
        return
    add_textbox_runs(
        root,
        x=1.15,
        y=2.28,
        w=10.30,
        h=1.45,
        runs=[
            (f"{piece.num}", "方正小标宋_GBK", title_size, GREEN, True),
            (display_title, "方正小标宋_GBK", title_size, BLACK, True),
        ],
        name="篇目标题",
        align="dist",
    )
    if piece.author:
        add_textbox_runs(
            root,
            x=10.08,
            y=2.02,
            w=2.30,
            h=0.70,
            runs=[(f"「{piece.author}」", "黑体", 28, "00B0F0", True)],
            name="考频",
            align="ctr",
        )
    add_textbox_runs(
        root,
        x=0.40,
        y=6.43,
        w=2.30,
        h=0.40,
        runs=[(piece.volume, "微软雅黑", 20, WHITE, False)],
        name="篇目标题册别",
        align="l",
    )


def first_tif_name(lines: list[wenyan.StyledLine]) -> str | None:
    for line in lines:
        match = re.search(r'data-fbd-tif="([^"]+)"', line.text)
        if match:
            return Path(match.group(1)).name
        match = re.search(r'<img[^>]+src="([^"]+)"', line.text)
        if match:
            return Path(match.group(1)).name
    return None


def extract_tif_names_from_text(text: str) -> list[str]:
    names: list[str] = []
    for match in re.finditer(r'data-fbd-tif="([^"]+)"', text, flags=re.I):
        names.append(Path(html.unescape(match.group(1))).name)
    if names:
        return names
    for match in re.finditer(r'<img\b[^>]+src="([^"]+)"', text, flags=re.I):
        names.append(Path(html.unescape(match.group(1))).name)
    return names


def index_groups(indices: object) -> list[tuple[int, int]]:
    groups: list[tuple[int, int]] = []
    values = [int(value) for value in indices]
    if not values:
        return groups
    start = previous = values[0]
    for value in values[1:]:
        if value <= previous + 1:
            previous = value
            continue
        groups.append((start, previous))
        start = previous = value
    groups.append((start, previous))
    return groups


def crop_practice_tif_image(image: object) -> tuple[object, int]:
    """Crop unused bottom Tianzige rows from 易错字词强化练习 TIF images."""

    try:
        import numpy as np
        from PIL import Image
    except Exception:
        return image, 0

    rgb = image.convert("RGB")
    arr = np.asarray(rgb)
    height, width = arr.shape[:2]
    if width < 1000 or height < 300:
        return rgb, 0

    channel = arr.astype("int16")
    strict_black = (channel[:, :, 0] < 120) & (channel[:, :, 1] < 120) & (channel[:, :, 2] < 120)
    dark_or_gray = (channel[:, :, 0] < 190) & (channel[:, :, 1] < 190) & (channel[:, :, 2] < 190)
    cyan_ink = (
        (channel[:, :, 0] < 170)
        & (channel[:, :, 1] > 80)
        & (channel[:, :, 2] > 80)
        & ((channel[:, :, 2] - channel[:, :, 0]) > 20)
    )
    nonwhite = (channel[:, :, 0] < 245) | (channel[:, :, 1] < 245) | (channel[:, :, 2] < 245)

    solid_rows = index_groups(np.where(strict_black.sum(axis=1) > width * 0.70)[0])
    row_bands: list[tuple[int, int]] = []
    for index in range(0, len(solid_rows) - 1, 2):
        top = solid_rows[index][0]
        bottom = solid_rows[index + 1][1]
        if bottom - top > 80:
            row_bands.append((top, bottom))
    if not row_bands:
        return rgb, 0

    grid_rows = index_groups(np.where(nonwhite.sum(axis=1) > width * 0.25)[0])
    grid_cols = index_groups(np.where(strict_black.sum(axis=0) > height * 0.20)[0])
    ink = (dark_or_gray | cyan_ink).copy()
    for top, bottom in grid_rows:
        ink[max(0, top - 3) : min(height, bottom + 4), :] = False
    for left, right in grid_cols:
        ink[:, max(0, left - 3) : min(width, right + 4)] = False

    content_row_indexes: list[int] = []
    for band_index, (top, bottom) in enumerate(row_bands):
        band_area = (bottom - top + 1) * width
        threshold = max(600, int(band_area * 0.00045))
        if int(ink[top : bottom + 1, :].sum()) > threshold:
            content_row_indexes.append(band_index)
    if not content_row_indexes:
        return rgb, 0

    last_content_row = max(content_row_indexes)
    crop_top = max(0, row_bands[0][0] - 1)
    crop_bottom = min(height, row_bands[last_content_row][1] + 1)
    cropped = rgb.crop((0, crop_top, width, crop_bottom))

    row_count = last_content_row + 1
    pad_ratio = min(0.16, max(0.02, 0.28 / max(1, row_count)))
    pad_total = int(round(cropped.height * pad_ratio))
    if pad_total > 0:
        top_pad = int(round(pad_total * 0.46))
        bottom_pad = pad_total - top_pad
        padded = Image.new("RGB", (cropped.width, cropped.height + top_pad + bottom_pad), "white")
        padded.paste(cropped, (0, top_pad))
        cropped = padded

    return cropped, row_count


def bibei_tif_png_info(tif_path: Path) -> tuple[bytes, int, int, int]:
    from PIL import Image

    with Image.open(tif_path) as image:
        converted = image.convert("RGB")
        row_count = 0
        if re.match(r"^T\d+\.tiff?$", tif_path.name, flags=re.I):
            converted, row_count = crop_practice_tif_image(converted)
        if converted.width > BIBEI_TIF_MAX_WIDTH_PX:
            new_height = max(1, round(converted.height * BIBEI_TIF_MAX_WIDTH_PX / converted.width))
            converted = converted.resize((BIBEI_TIF_MAX_WIDTH_PX, new_height), Image.Resampling.LANCZOS)
        buffer = io.BytesIO()
        converted.save(buffer, format="PNG", optimize=True)
        return buffer.getvalue(), converted.width, converted.height, row_count


def append_fixed_tif_picture(
    root: ET.Element,
    rels_root: ET.Element,
    media_assets: dict[str, bytes],
    tif_path: Path,
    *,
    x: float,
    y: float,
    w: float,
    h: float,
    name: str,
) -> int:
    package_path = wenyan.tif_package_path(tif_path)
    if package_path not in media_assets:
        media_assets[package_path] = wenyan.image_to_png_bytes(tif_path)
    rel_id = wenyan.add_image_relationship(rels_root, wenyan.posixpath.relpath(package_path, "ppt/slides"))
    return wenyan.append_picture(root, rel_id=rel_id, x=x, y=y, w=w, h=h, name=name)


def append_practice_tif_picture(
    root: ET.Element,
    rels_root: ET.Element,
    media_assets: dict[str, bytes],
    tif_path: Path,
    *,
    name: str,
) -> int:
    package_path = wenyan.tif_package_path(tif_path)
    data, pixel_w, pixel_h, row_count = bibei_tif_png_info(tif_path)
    if package_path not in media_assets:
        media_assets[package_path] = data
    rel_id = wenyan.add_image_relationship(rels_root, wenyan.posixpath.relpath(package_path, "ppt/slides"))

    aspect = pixel_w / pixel_h if pixel_h else 1.0
    picture_w = min(BIBEI_TIF_MAX_W_IN, BIBEI_TIF_MAX_H_IN * aspect)
    picture_h = picture_w / aspect if aspect else BIBEI_TIF_MAX_H_IN
    picture_x = (wenyan.SLIDE_W_IN - picture_w) / 2
    if picture_h >= 4.8:
        picture_y = BIBEI_TIF_TITLE_TALL_Y
    elif picture_h >= 3.0:
        picture_y = BIBEI_TIF_TITLE_MID_Y
    else:
        picture_y = BIBEI_TIF_TITLE_SHORT_Y
    if picture_y + picture_h > BIBEI_TIF_FOOTER_TOP_IN:
        picture_y = max(1.20, BIBEI_TIF_FOOTER_TOP_IN - picture_h)

    return wenyan.append_picture(
        root,
        rel_id=rel_id,
        x=picture_x,
        y=picture_y,
        w=picture_w,
        h=picture_h,
        name=f"{name} {row_count}行",
    )


def append_open_topic_picture_row(
    root: ET.Element,
    rels_root: ET.Element,
    media_assets: dict[str, bytes],
    *,
    tif_names: list[str],
    labels: list[str],
    tif_search_dirs: list[Path],
    x: float,
    y: float,
    w: float,
) -> float:
    images: list[tuple[str, Path, bytes, int, int]] = []
    for tif_name in tif_names:
        tif_path = wenyan.resolve_tif_path(tif_name, tif_search_dirs)
        if not tif_path:
            continue
        data, pixel_w, pixel_h, _row_count = bibei_tif_png_info(tif_path)
        images.append((tif_name, tif_path, data, pixel_w, pixel_h))
    if not images:
        return 0.0

    image_y = y + OPEN_TOPIC_IMAGE_TOP_PAD_IN
    gap = OPEN_TOPIC_IMAGE_GAP_IN
    max_h = OPEN_TOPIC_IMAGE_MAX_H_IN
    total_aspect = sum((pixel_w / pixel_h) if pixel_h else 1.0 for _name, _path, _data, pixel_w, pixel_h in images)
    available_w = max(1.0, w - gap * (len(images) - 1))
    picture_h = min(max_h, available_w / max(0.1, total_aspect))
    picture_sizes: list[tuple[float, float]] = []
    for _name, _path, _data, pixel_w, pixel_h in images:
        aspect = (pixel_w / pixel_h) if pixel_h else 1.0
        picture_sizes.append((picture_h * aspect, picture_h))
    total_w = sum(size[0] for size in picture_sizes) + gap * (len(picture_sizes) - 1)
    cursor_x = x + max(0.0, (w - total_w) / 2)
    centers: list[float] = []
    for (tif_name, tif_path, data, _pixel_w, _pixel_h), (picture_w, picture_h) in zip(images, picture_sizes):
        package_path = wenyan.tif_package_path(tif_path)
        if package_path not in media_assets:
            media_assets[package_path] = data
        rel_id = wenyan.add_image_relationship(rels_root, wenyan.posixpath.relpath(package_path, "ppt/slides"))
        wenyan.append_picture(
            root,
            rel_id=rel_id,
            x=cursor_x,
            y=image_y,
            w=picture_w,
            h=picture_h,
            name=f"open topic image {tif_name}",
        )
        centers.append(cursor_x + picture_w / 2)
        cursor_x += picture_w + gap

    label_y = image_y + picture_h + 0.04
    for index, center_x in enumerate(centers):
        label = labels[index] if index < len(labels) else f"图{index + 1}"
        wenyan.append_textbox(
            root,
            x=center_x - 0.35,
            y=label_y,
            w=0.70,
            h=OPEN_TOPIC_IMAGE_LABEL_H_IN,
            lines=[wenyan.StyledLine(label, "open_image_labels")],
            name=f"open topic image label {label}",
        )
    return (
        OPEN_TOPIC_IMAGE_TOP_PAD_IN
        + picture_h
        + OPEN_TOPIC_IMAGE_LABEL_H_IN
        + OPEN_TOPIC_IMAGE_BOTTOM_PAD_IN
    )


def bibei_visual_line_count(line: wenyan.StyledLine, width_in: float) -> int:
    role = line.role or "body"
    if role == "open_image_row":
        return 4
    if role == "open_image_labels":
        return 1
    capacity = wenyan.line_capacity_for_width(width_in)
    units = 0.0
    max_visual_line = 0
    for inline in wenyan.parse_inline_runs(line.text):
        if inline.tif_name:
            units += 1.0
        else:
            units += wenyan.text_units(re.sub(r"<[^>]+>", "", html.unescape(inline.text)))
        max_visual_line = max(max_visual_line, int(units // capacity))
    if role in {"heading", "module_header", "module", "subheading", "center_heading", "rush_day", "rush_group"}:
        return 1
    return max(1, max_visual_line + 1)


def bibei_line_height(line: wenyan.StyledLine) -> float:
    role = line.role or "body"
    base = wenyan.line_height_for_role(role)
    if role in {"heading", "subheading", "rush_day", "rush_group"}:
        return max(0.42, base)
    return max(0.50, base)


def bibei_inline_answer_count(line: wenyan.StyledLine) -> int:
    return sum(1 for inline in wenyan.parse_inline_runs(line.text) if inline.answer_fill)


def bibei_answer_display_text(text: str) -> str:
    plain = wenyan.strip_inline_tags(html.unescape(text))
    plain = re.sub(r"[_＿]{2,}", " ", plain)
    plain = plain.replace("（", "").replace("(", "").replace("）", "").replace(")", "")
    return re.sub(r"\s+", " ", plain).strip(" \t\r\n　")


def bibei_overlay_box_size(answer_text: str, x: float, y: float, w: float, line_h: float) -> tuple[float, float]:
    units = max(1.0, wenyan.text_units(answer_text))
    char_w = 0.24
    target_w = min(w, max(0.34, units * char_w + 0.12))
    line_capacity = max(4.0, target_w / char_w)
    line_count = max(1, math.ceil(units / line_capacity))
    return target_w, max(0.34, line_count * line_h + 0.06)


def bibei_overlay_role(role: str) -> str:
    if role in {"analysis", "segmentation_analysis"}:
        return role
    return "answer"


def append_bibei_overlay_answer_shape(
    root: ET.Element,
    *,
    x: float,
    y: float,
    w: float,
    h: float,
    text: str,
    role: str,
    name: str,
) -> int:
    return wenyan.append_textbox(
        root,
        x=x,
        y=y,
        w=w,
        h=h,
        lines=[wenyan.StyledLine(text, role)],
        name=name,
    )


def append_bibei_answer_overlays(
    root: ET.Element,
    *,
    x: float,
    y: float,
    w: float,
    lines: list[wenyan.StyledLine],
    name: str,
    inline_y_offset: float = 0.0,
) -> list[wenyan.AnimationTarget]:
    animations: list[wenyan.AnimationTarget] = []
    cursor = 0.0
    index = 0
    while index < len(lines):
        line = lines[index]
        role = line.role or "body"
        line_h = bibei_line_height(line)
        visual_lines = bibei_visual_line_count(line, w)

        if role == "long_answer":
            group: list[wenyan.StyledLine] = []
            group_h = 0.0
            while index < len(lines) and (lines[index].role or "body") == "long_answer":
                group_line = lines[index]
                group.append(group_line)
                group_h += bibei_line_height(group_line) * bibei_visual_line_count(group_line, w)
                index += 1
            for group_line in group:
                display = bibei_answer_display_text(group_line.text)
                if not display:
                    cursor += bibei_line_height(group_line) * bibei_visual_line_count(group_line, w)
                    continue
                box_w, box_h = bibei_overlay_box_size(display, x, y + cursor, w, bibei_line_height(group_line))
                shape_id = append_bibei_overlay_answer_shape(
                    root,
                    x=x,
                    y=y + cursor,
                    w=box_w,
                    h=box_h,
                    text=display,
                    role=bibei_overlay_role(role),
                    name=f"{name} answer {len(animations) + 1}",
                )
                animations.append(wenyan.AnimationTarget(shape_id=shape_id, effect="appear", text_target=False))
            cursor += group_h
            continue

        if role in wenyan.ANIMATED_LINE_ROLES:
            display = bibei_answer_display_text(line.text)
            if not display:
                cursor += line_h * visual_lines
                index += 1
                continue
            box_w, box_h = bibei_overlay_box_size(display, x, y + cursor, w, line_h)
            shape_id = append_bibei_overlay_answer_shape(
                root,
                x=x,
                y=y + cursor,
                w=box_w,
                h=box_h,
                text=display,
                role=bibei_overlay_role(role),
                name=f"{name} answer {len(animations) + 1}",
            )
            animations.append(wenyan.AnimationTarget(shape_id=shape_id, effect="appear", text_target=False))
        else:
            inline_units = 0.0
            for inline in wenyan.parse_inline_runs(line.text):
                inline_text = html.unescape(wenyan.strip_inline_tags(inline.text))
                inline_units_text = wenyan.text_units(inline_text)
                if not inline.answer_fill:
                    inline_units += inline_units_text + (1.0 if inline.tif_name else 0.0)
                    continue
                display = bibei_answer_display_text(inline.text)
                if not display:
                    inline_units += inline_units_text
                    continue
                leading_trim = len(inline.text) - len(inline.text.lstrip(" \t\r\n　_＿"))
                leading_text = inline.text[:leading_trim]
                leading_units = wenyan.text_units(wenyan.strip_inline_tags(html.unescape(leading_text)))
                start_units = inline_units + leading_units
                char_w = 0.24
                start_line = int(start_units // max(1.0, wenyan.line_capacity_for_width(w)))
                x_offset = (start_units % max(1.0, wenyan.line_capacity_for_width(w))) * char_w
                box_x = x + x_offset
                box_w, box_h = bibei_overlay_box_size(display, box_x, y + cursor + start_line * line_h, max(0.42, w - x_offset), line_h)
                shape_id = append_bibei_overlay_answer_shape(
                    root,
                    x=box_x,
                    y=y + cursor + start_line * line_h + inline_y_offset,
                    w=box_w,
                    h=box_h,
                    text=display,
                    role=bibei_overlay_role(role),
                    name=f"{name} answer {len(animations) + 1}",
                )
                animations.append(wenyan.AnimationTarget(shape_id=shape_id, effect="appear", text_target=False))
                inline_units += inline_units_text
        cursor += line_h * visual_lines
        index += 1
    return animations


def append_bibei_base_textboxes(
    root: ET.Element,
    *,
    x: float,
    y: float,
    w: float,
    lines: list[wenyan.StyledLine],
    name: str,
) -> None:
    cursor = 0.0
    for index, line in enumerate(lines, 1):
        line_h = bibei_line_height(line)
        visual_lines = bibei_visual_line_count(line, w)
        wenyan.append_textbox(
            root,
            x=x,
            y=y + cursor,
            w=w,
            h=max(0.34, line_h * visual_lines + 0.06),
            lines=[line],
            name=f"{name} line {index}",
            answer_mode="base",
        )
        cursor += line_h * visual_lines


def add_bibei_flow_content(
    root: ET.Element,
    rels_root: ET.Element,
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
    animations = append_bibei_answer_overlays(
        root,
        x=base_x,
        y=y,
        w=base_w,
        lines=content_lines,
        name=f"{piece.title} content",
    )
    wenyan.set_slide_animations(root, animations)


def add_open_topic_content(
    root: ET.Element,
    rels_root: ET.Element,
    piece: wenyan.Piece,
    chunk: list[wenyan.StyledLine],
    *,
    media_assets: dict[str, bytes],
    tif_search_dirs: list[Path],
    content_box: tuple[float, float, float, float],
) -> None:
    wenyan.clear_text_shapes(root, keep_substrings=("能力",))
    base_x, base_y, base_w, base_h = content_box
    cursor = 0.0
    animations: list[wenyan.AnimationTarget] = []
    index = 0
    while index < len(chunk):
        line = chunk[index]
        role = line.role or "body"
        if role == "open_image_row":
            labels: list[str] = []
            if index + 1 < len(chunk) and (chunk[index + 1].role or "body") == "open_image_labels":
                labels = open_figure_labels(chunk[index + 1].text) or chunk[index + 1].text.split()
                index += 1
            tif_names = line.meta.get("tif_names", []) if line.meta else []
            if not tif_names:
                tif_names = extract_tif_names_from_text(line.text)
            used_h = append_open_topic_picture_row(
                root,
                rels_root,
                media_assets,
                tif_names=[str(name) for name in tif_names],
                labels=labels,
                tif_search_dirs=tif_search_dirs,
                x=base_x + 0.18,
                y=base_y + cursor + 0.04,
                w=base_w - 0.36,
            )
            cursor += max(used_h, 0.65)
            index += 1
            continue
        if role == "open_image_labels":
            index += 1
            continue

        line_h = bibei_line_height(line)
        visual_lines = bibei_visual_line_count(line, base_w)
        shape_h = max(0.34, line_h * visual_lines + 0.04)
        wenyan.append_textbox(
            root,
            x=base_x,
            y=base_y + cursor,
            w=base_w,
            h=shape_h,
            lines=[line],
            name=f"{piece.title} open line {index + 1}",
            answer_mode="base",
        )
        animations.extend(
            append_bibei_answer_overlays(
                root,
                x=base_x,
                y=base_y + cursor,
                w=base_w,
                lines=[line],
                name=f"{piece.title} open line {index + 1}",
                inline_y_offset=OPEN_TOPIC_INLINE_ANSWER_Y_OFFSET_IN,
            )
        )
        cursor += line_h * visual_lines
        index += 1
    wenyan.set_slide_animations(root, animations)


def add_bibei_content(
    root: ET.Element,
    rels_root: ET.Element,
    piece: wenyan.Piece,
    chunk: list[wenyan.StyledLine],
    *,
    media_assets: dict[str, bytes],
    tif_search_dirs: list[Path],
    content_box: tuple[float, float, float, float],
) -> None:
    if piece.volume == "册子" and title_plain(piece.title).startswith("专题"):
        add_open_topic_content(
            root,
            rels_root,
            piece,
            chunk,
            media_assets=media_assets,
            tif_search_dirs=tif_search_dirs,
            content_box=content_box,
        )
        return

    tif_name = first_tif_name(chunk)
    if tif_name and re.match(r"^T\d+\.tiff?$", tif_name, flags=re.I):
        wenyan.clear_text_shapes(root, keep_substrings=("能力",))
        heading = next((line for line in chunk if re.match(r"^[一二三四五六七八九十]+、", line.text)), None)
        if heading:
            wenyan.append_textbox(root, x=0.70, y=0.72, w=6.9, h=0.55, lines=[heading], name=f"{piece.title} 易错字标题")
        tif_path = wenyan.resolve_tif_path(tif_name, tif_search_dirs)
        if tif_path:
            append_practice_tif_picture(
                root,
                rels_root,
                media_assets,
                tif_path,
                name=f"{piece.title} 易错字图",
            )
        wenyan.set_slide_animations(root, [])
        return

    wenyan.add_content(
        root,
        rels_root,
        piece,
        chunk,
        media_assets=media_assets,
        tif_search_dirs=tif_search_dirs,
        content_box=content_box,
    )


@isolated_generation
def build_bibei_deck(
    template: Path,
    output: Path,
    pieces: list[wenyan.Piece],
    answer_map: dict[tuple[str, int], wenyan.Piece],
    *,
    volume: str,
    start: int,
    end: int,
    tif_search_dirs: list[Path],
    include_training_label: bool = True,
    cover_subtitle: str | None = None,
    directory_page_size: int = 6,
) -> None:
    wenyan.PARAGRAPH_LINE_SPACING_PCT = 130000
    wenyan.paragraph_properties = bibei_paragraph_properties
    wenyan.role_style = bibei_role_style
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(template, "r") as zin:
        slide_count_template = len([n for n in zin.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)])
        title_src, content_src, final_sources, _template_content_box = wenyan.template_contract(slide_count_template)
        content_box = BOOKLET_CONTENT_BOX if volume == "册子" else BIBEI_CONTENT_BOX
        generated_slides: list[tuple[int, bytes, bytes]] = []
        extra_media_assets: dict[str, bytes] = {}
        piece_chunks: list[tuple[wenyan.Piece, list[list[wenyan.StyledLine]]]] = []
        title_slide_numbers: dict[tuple[str, int], int] = {}
        directory_page_size = max(1, directory_page_size)
        directory_pages = max(1, (len(pieces) + directory_page_size - 1) // directory_page_size)
        next_generated_slide = 3 + directory_pages
        for piece in pieces:
            chunks = make_bibei_piece_slide_chunks(piece, answer_map.get(wenyan.piece_key(piece)))
            piece_chunks.append((piece, chunks))
            title_slide_numbers[wenyan.piece_key(piece)] = next_generated_slide
            next_generated_slide += 1 + len(chunks)

        def add_from(src: int, edit_func=lambda root, rels_root: None) -> None:
            xml, rels = wenyan.make_slide_parts(zin, src, edit_func)
            generated_slides.append((src, xml, rels))

        add_from(
            1,
            lambda root, rels_root: add_bibei_cover_content(
                root,
                volume,
                start,
                end,
                subtitle=cover_subtitle,
                include_training_label=include_training_label,
            ),
        )
        add_from(2)
        for page_index in range(directory_pages):
            directory_pieces = pieces[page_index * directory_page_size : page_index * directory_page_size + directory_page_size]
            directory_show_training = include_training_label and page_index == directory_pages - 1
            add_from(
                3,
                lambda root, rels_root, directory_pieces=directory_pieces, directory_show_training=directory_show_training: add_bibei_directory_content(
                    root,
                    rels_root,
                    directory_pieces,
                    title_slide_numbers,
                    include_training_label=directory_show_training,
                ),
            )
        for piece, chunks in piece_chunks:
            add_from(title_src, lambda root, rels_root, piece=piece: add_bibei_piece_title(root, piece))
            for chunk in chunks:
                add_from(
                    content_src,
                    lambda root, rels_root, piece=piece, chunk=chunk: add_bibei_content(
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
        presentation_rels = wenyan.update_presentation_rels(wenyan.parse_xml(zin.read("ppt/_rels/presentation.xml.rels")), slide_total)

        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            skip = {"[Content_Types].xml", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"}
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
            for package_path, data in extra_media_assets.items():
                if package_path not in written_media:
                    zout.writestr(package_path, data)
                    written_media.add(package_path)
            zout.writestr("[Content_Types].xml", content_types)
            zout.writestr("ppt/presentation.xml", presentation_xml)
            zout.writestr("ppt/_rels/presentation.xml.rels", presentation_rels)
            for i, (_src, xml, rels) in enumerate(generated_slides, 1):
                zout.writestr(f"ppt/slides/slide{i}.xml", xml)
                zout.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", rels)
    try:
        wenyan.postprocess_tif_positions_with_powerpoint(output, adjust_answer_overlays=True)
    except Exception as exc:
        print(f"提示：PPTX 后处理失败，已保留 XML 生成结果：{exc}", file=sys.stderr)


def find_md_files(run_dir: Path) -> tuple[Path, Path]:
    body = sorted(run_dir.glob("正文_*.md"))
    answer = sorted(run_dir.glob("答案_*.md"))
    if not body or not answer:
        raise FileNotFoundError(f"未在目录中找到正文_*.md 和 答案_*.md：{run_dir}")
    return body[0], answer[0]


def generate_bibei_test_pptx(
    *,
    md_run: Path,
    template: Path,
    output_dir: Path,
    limit_pieces: int = 6,
    progress: Callable[[str], None] | None = None,
) -> tuple[Path, ParseStats]:
    def emit(message: str) -> None:
        if progress:
            progress(message)
        else:
            print(message)

    body_md, answer_md = find_md_files(md_run)
    body_pieces = parse_body_pieces(body_md)
    answer_pieces = parse_answer_pieces(answer_md)
    answer_map = build_bibei_answer_piece_map(body_pieces, answer_pieces)
    tif_search_dirs = wenyan.infer_tif_search_dirs(md_run)

    selected = body_pieces[:limit_pieces]
    if not selected:
        raise ValueError("正文 MD 中没有解析到篇目")
    output = output_dir / FIRST_GROUP_NAME
    emit(f"正文篇目数：{len(body_pieces)}")
    emit(f"答案篇目数：{len(answer_pieces)}")
    emit(f"生成测试 PPTX：{output}")
    build_bibei_deck(
        template,
        output,
        selected,
        answer_map,
        volume=selected[0].volume,
        start=selected[0].num,
        end=selected[-1].num,
        tif_search_dirs=tif_search_dirs,
    )
    manifest = {
        "md_run": str(md_run),
        "template": str(template),
        "output": str(output),
        "pieces": [f"{piece.num}.{piece.title}" for piece in selected],
        "body_count": len(body_pieces),
        "answer_count": len(answer_pieces),
    }
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "test_bibei_generation_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return output, ParseStats(len(body_pieces), len(answer_pieces), [p.title for p in selected])


def generate_bibei_main_pptx_all(
    *,
    md_run: Path,
    template: Path,
    output_dir: Path,
    progress: Callable[[str], None] | None = None,
) -> tuple[list[Path], ParseStats]:
    def emit(message: str) -> None:
        if progress:
            progress(message)
        else:
            print(message)

    body_md, answer_md = find_md_files(md_run)
    body_pieces = parse_body_pieces(body_md)
    answer_pieces = parse_answer_pieces(answer_md)
    answer_map = build_bibei_answer_piece_map(body_pieces, answer_pieces)
    piece_by_num = {piece.num: piece for piece in body_pieces}
    tif_search_dirs = wenyan.infer_tif_search_dirs(md_run)
    output_dir.mkdir(parents=True, exist_ok=True)

    outputs: list[Path] = []
    for index, start, end in BIBEI_MAIN_GROUP_SPECS:
        selected = [piece_by_num[num] for num in range(start, end + 1) if num in piece_by_num]
        missing = [num for num in range(start, end + 1) if num not in piece_by_num]
        if missing:
            raise ValueError(f"正文 MD 缺少篇目：{missing}")
        output = output_dir / f"{index} 篇目{start}-{end}+「6+1训练赢」.pptx"
        emit(f"生成：{output.name}")
        build_bibei_deck(
            template,
            output,
            selected,
            answer_map,
            volume=selected[0].volume,
            start=start,
            end=end,
            tif_search_dirs=tif_search_dirs,
        )
        outputs.append(output)

    manifest = {
        "md_run": str(md_run),
        "template": str(template),
        "output_dir": str(output_dir),
        "outputs": [str(path) for path in outputs],
        "pieces": [f"{piece.num}.{piece.title}" for piece in body_pieces],
        "body_count": len(body_pieces),
        "answer_count": len(answer_pieces),
    }
    (output_dir / "bibei_main_generation_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return outputs, ParseStats(len(body_pieces), len(answer_pieces), [p.title for p in body_pieces[:6]])


def marker_plain(line: str) -> str:
    return re.sub(r"\s+", "", title_plain(line))


def find_marker_index(lines: list[str], marker: str, *, start: int = 0) -> int:
    compact_marker = re.sub(r"\s+", "", marker)
    for index in range(start, len(lines)):
        if marker_plain(lines[index]) == compact_marker:
            return index
    return -1


def parse_main_extra_body_pieces(md_path: Path) -> list[wenyan.Piece]:
    lines = [line for line in wenyan.clean_lines(md_path.read_text(encoding="utf-8")) if line.strip()]
    start = find_marker_index(lines, "附赠选修12篇")
    end = find_marker_index(lines, "150组冲刺练", start=start + 1 if start >= 0 else 0)
    if start < 0 or end < 0:
        return []
    source = lines[start + 1 : end]
    starts: list[tuple[int, int]] = []
    for index, line in enumerate(source):
        value = plain(line)
        if not re.fullmatch(r"6[1-9]|7[0-2]", value):
            continue
        lookahead = [plain(item) for item in nonempty_after(source, index, 4)]
        if len(lookahead) >= 3 and lookahead[1].isdigit() and "理解性默写" in lookahead[2]:
            starts.append((index, int(value)))

    pieces: list[wenyan.Piece] = []
    for pos, (piece_start, num) in enumerate(starts):
        piece_end = starts[pos + 1][0] if pos + 1 < len(starts) else len(source)
        block = [line for line in source[piece_start:piece_end] if line.strip()]
        if len(block) < 4:
            continue
        title = title_plain(block[1])
        content_start = 3
        while content_start < len(block) and "理解性默写" not in title_plain(block[content_start]):
            content_start += 1
        content = ["一、理解性默写。", *[normalize_body_numbering(line) for line in block[content_start + 1 :]]]
        pieces.append(wenyan.Piece(volume="附赠选修12篇", num=num, title=title, author="", lines=[title, *content]))
    return pieces


def parse_main_extra_answer_pieces(md_path: Path, body_pieces: list[wenyan.Piece]) -> list[wenyan.Piece]:
    body_titles = {piece.num: piece.title for piece in body_pieces}
    lines = [line for line in wenyan.clean_lines(md_path.read_text(encoding="utf-8")) if line.strip()]
    start = find_marker_index(lines, "附赠选修12篇")
    end = find_marker_index(lines, "150组冲刺练", start=start + 1 if start >= 0 else 0)
    if start < 0 or end < 0:
        return []
    source = lines[start + 1 : end]
    starts: list[tuple[int, int]] = []
    for index, line in enumerate(source):
        match = re.match(r"^(6[1-9]|7[0-2])\.\s*(.+)$", title_plain(line))
        if match:
            starts.append((index, int(match.group(1))))

    pieces: list[wenyan.Piece] = []
    for pos, (piece_start, num) in enumerate(starts):
        piece_end = starts[pos + 1][0] if pos + 1 < len(starts) else len(source)
        block = [line for line in source[piece_start:piece_end] if line.strip()]
        if not block:
            continue
        answer_start = 1
        while answer_start < len(block) and not wenyan.answer_item_match(block[answer_start]):
            answer_start += 1
        title = body_titles.get(num) or re.sub(r"^\d+\.\s*", "", title_plain(block[0]))
        pieces.append(
            wenyan.Piece(
                volume="附赠选修12篇",
                num=num,
                title=title,
                author="",
                lines=[title, "一、理解性默写。", *[normalize_answer_line(line) for line in block[answer_start:]]],
            )
        )
    return pieces


def parse_main_rush_piece(md_path: Path) -> wenyan.Piece | None:
    lines = [line for line in wenyan.clean_lines(md_path.read_text(encoding="utf-8")) if line.strip()]
    start = find_marker_index(lines, "150组冲刺练")
    if start < 0:
        return None
    body_lines = [
        normalize_body_numbering(line)
        for line in lines[start + 1 :]
        if marker_plain(line) != "150组冲刺练" and not plain(line).isdigit()
    ]
    return wenyan.Piece(volume="主书", num=12, title="150组冲刺练", author="", lines=["150组冲刺练", "一、150组冲刺练", *body_lines])


def parse_main_rush_answer_piece(md_path: Path) -> wenyan.Piece | None:
    lines = [line for line in wenyan.clean_lines(md_path.read_text(encoding="utf-8")) if line.strip()]
    start = find_marker_index(lines, "150组冲刺练")
    end = find_marker_index(lines, "《默写必背突围手册》答案解析", start=start + 1 if start >= 0 else 0)
    if start < 0 or end < 0:
        return None
    answer_lines = [normalize_answer_line(line) for line in lines[start + 1 : end] if marker_plain(line) != "150组冲刺练"]
    return wenyan.Piece(volume="主书", num=12, title="150组冲刺练", author="", lines=["150组冲刺练", "一、150组冲刺练", *answer_lines])


BIBEI_BOOKLET_VOLUME_LABELS = [
    ("必修上册", "1 分册强化练必修上册.pptx"),
    ("必修下册", "2 分册强化练必修下册.pptx"),
    ("选择性必修上册", "3 分册强化练选必修上册.pptx"),
    ("选择性必修中册", "4 分册强化练选必修中册.pptx"),
    ("选择性必修下册", "5 分册强化练选必修下册.pptx"),
    ("教材外必背篇目", "6 分册强化练教材外必背篇目.pptx"),
]


def find_pocket_md_file(run_dir: Path) -> Path:
    candidates = sorted(run_dir.glob("口袋本_*.md"))
    if not candidates:
        candidates = sorted(run_dir.glob("*.md"))
    if not candidates:
        raise FileNotFoundError(f"未在口袋本目录中找到 md：{run_dir}")
    return candidates[0]


def canonical_booklet_title(title: str) -> str:
    value = title_plain(title)
    value = re.sub(r"^\d{1,3}[\.、]\s*", "", value)
    value = value.replace("（", "(").replace("）", ")")
    value = re.sub(r"[《》〈〉\s　·．.、]", "", value)
    value = value.replace("赤壁怀古", "赤壁怀古")
    return value


def replace_tif_tags_with_text(text: str) -> str:
    def repl(match: re.Match[str]) -> str:
        tag = match.group(0)
        name_match = re.search(r'data-fbd-tif="([^"]+)"', tag) or re.search(r'src="([^"]+)"', tag)
        if not name_match:
            return ""
        tif_name = Path(html.unescape(name_match.group(1))).name.lower()
        return BIBEI_TIF_TEXT_ALIASES.get(tif_name, "")

    return re.sub(r"<img\b[^>]*>", repl, text, flags=re.I)


PINYIN_CHARS = "A-Za-zＡ-Ｚａ-ｚāáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜüÜɑɡńňǹḿń"


def remove_pinyin_notes(text: str) -> str:
    return re.sub(rf"[（(]\s*[{PINYIN_CHARS}\s]+\s*[）)]", "", text)


def normalize_moxie_match_text(text: str) -> str:
    text = replace_tif_tags_with_text(text)
    text = wenyan.strip_inline_tags(html.unescape(text))
    text = remove_pinyin_notes(text)
    text = re.sub(r"[0-9０-９]+", "", text)
    text = re.sub(r"[\s　，。！？；：、,.!?;:“”\"‘’（）()《》〈〉·—\-…⋯]", "", text)
    return text


def display_moxie_answer(text: str) -> str:
    text = remove_pinyin_notes(text)
    text = re.sub(r"\s+", "", text)
    return text


def is_booklet_moxie_source_note(line: str) -> bool:
    value = title_plain(line)
    return bool(re.match(r"^(?:诗|词|曲|文言文)[·・]\S+", value))


def is_booklet_moxie_match_noise_line(line: str) -> bool:
    value = title_plain(line)
    compact = re.sub(r"[_＿\s\u3000]+", "", value)
    if not compact:
        return True
    if re.fullmatch(r"[.。…⋯]+", compact):
        return True
    return is_booklet_moxie_source_note(line)


def parse_pocket_poem_index(pocket_md: Path | None) -> dict[str, str]:
    if pocket_md is None or not pocket_md.exists():
        return {}
    lines = wenyan.clean_lines(pocket_md.read_text(encoding="utf-8"))
    starts: list[tuple[int, int, str]] = []
    for index, line in enumerate(lines):
        match = re.match(r"^(\d{1,3})\.\s*(.+)$", title_plain(line))
        if match and 1 <= int(match.group(1)) <= 72:
            starts.append((index, int(match.group(1)), match.group(2).strip()))

    poem_index: dict[str, str] = {}
    for pos, (start, _num, title) in enumerate(starts):
        end = starts[pos + 1][0] if pos + 1 < len(starts) else len(lines)
        block_text = "\n".join(lines[start:end])
        spans = re.findall(r'<span\s+class="uline"[^>]*>(.*?)</span>', block_text, flags=re.S)
        source_text = "\n".join(spans) if spans else block_text
        normalized = normalize_moxie_match_text(source_text)
        if normalized:
            poem_index[canonical_booklet_title(title)] = normalized
    return poem_index


@dataclass(frozen=True)
class BookletBlank:
    kind: str


def booklet_line_tokens(line: str) -> list[str | BookletBlank]:
    tokens: list[str | BookletBlank] = []
    pos = 0
    pattern = re.compile(
        r'<span\s+class="uline"[^>]*>\s*[_＿]+\s*</span>|[（(]\s*[_＿]{2,}\s*[）)]|[_＿]{3,}'
    )
    for match in pattern.finditer(line):
        if match.start() > pos:
            tokens.append(line[pos : match.start()])
        matched = match.group(0)
        kind = "uline" if "class=\"uline\"" in matched or "class='uline'" in matched else "paren"
        if re.fullmatch(r"[_＿]{3,}", matched):
            kind = "plain"
        tokens.append(BookletBlank(kind))
        pos = match.end()
    if pos < len(line):
        tokens.append(line[pos:])
    return tokens


def allocate_moxie_gap(gap: str, blanks: list[BookletBlank]) -> list[str]:
    if not blanks:
        return []
    if not gap:
        return ["" for _ in blanks]
    if len(blanks) == 1:
        return [display_moxie_answer(gap)]

    result: list[str] = []
    rest = gap
    for index, blank in enumerate(blanks):
        remaining = len(blanks) - index - 1
        if blank.kind == "paren":
            take = max(0, min(1, len(rest) - remaining))
            if take == 0 and rest:
                take = 1
            result.append(display_moxie_answer(rest[:take]))
            rest = rest[take:]
        else:
            if remaining == 0:
                take = len(rest)
            else:
                take = max(0, len(rest) - remaining)
            result.append(display_moxie_answer(rest[:take]))
            rest = rest[take:]
    if rest and result:
        result[-1] = display_moxie_answer(result[-1] + rest)
    return result


def moxie_fixed_and_blank_groups(lines: list[str]) -> tuple[list[str], list[list[BookletBlank]], int]:
    fixed_parts: list[str] = []
    blank_groups: list[list[BookletBlank]] = [[]]
    current_text: list[str] = []
    blank_count = 0

    def flush_text() -> None:
        text = normalize_moxie_match_text("".join(current_text))
        current_text.clear()
        if text:
            fixed_parts.append(text)
            blank_groups.append([])

    for line in lines:
        for token in booklet_line_tokens(line):
            if isinstance(token, BookletBlank):
                flush_text()
                blank_groups[-1].append(token)
                blank_count += 1
            else:
                current_text.append(token)
        current_text.append("\n")
    flush_text()
    return fixed_parts, blank_groups, blank_count


def find_moxie_fixed_part(poem_text: str, fixed: str, start: int) -> tuple[int, int] | None:
    found = poem_text.find(fixed, start)
    if found >= 0:
        return found, len(fixed)

    probes: list[str] = []
    if len(fixed) > 12:
        probes.append(fixed[:12])
    if len(fixed) > 6:
        probes.append(fixed[:6])
    elif len(fixed) >= 4:
        probes.append(fixed)

    candidates: list[tuple[int, int]] = []
    for probe in probes:
        probe_found = poem_text.find(probe, start)
        if probe_found >= 0:
            candidates.append((probe_found, len(probe)))
        if len(probe) >= 4:
            for index in range(len(probe)):
                variant = probe[:index] + probe[index + 1 :]
                if len(variant) < 3:
                    continue
                variant_found = poem_text.find(variant, start)
                if variant_found >= 0:
                    candidates.append((variant_found, len(variant)))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (item[0], -item[1]))
    return candidates[0]


def derive_moxie_answers_from_poem(lines: list[str], poem_text: str) -> list[str]:
    match_lines = [line for line in lines if not is_booklet_moxie_match_noise_line(line)]
    fixed_parts, blank_groups, blank_count = moxie_fixed_and_blank_groups(match_lines)
    if not poem_text or blank_count == 0:
        return []
    answers: list[str] = []
    pos = 0
    for index, fixed in enumerate(fixed_parts):
        found_part = find_moxie_fixed_part(poem_text, fixed, pos)
        if found_part is None:
            return []
        found, matched_len = found_part
        gap = poem_text[pos:found]
        answers.extend(allocate_moxie_gap(gap, blank_groups[index]))
        pos = found + matched_len
    answers.extend(allocate_moxie_gap(poem_text[pos:], blank_groups[-1]))
    return answers


def fill_booklet_moxie_lines(lines: list[str], answers: list[str]) -> list[str]:
    if not answers:
        return lines
    answer_index = 0
    filled_lines: list[str] = []
    blank_pattern = re.compile(
        r'<span\s+class=["\']uline["\'][^>]*>\s*(?:<span\s+class=["\']blank["\'][^>]*>\s*</span>|[_＿]{2,})\s*</span>'
        r'|([\(\uff08])\s*(?:<span\s+class=["\']blank["\'][^>]*>\s*</span>|[_＿]{2,})\s*([\)\uff09])'
        r'|[_＿]{3,}',
        flags=re.I,
    )

    def fill_line(line: str) -> str:
        nonlocal answer_index

        def repl(match: re.Match[str]) -> str:
            nonlocal answer_index
            if answer_index >= len(answers):
                return match.group(0)
            answer = answers[answer_index]
            answer_index += 1
            matched = match.group(0)
            left_paren = match.group(1)
            right_paren = match.group(2)
            if left_paren is not None and right_paren is not None:
                return f"{left_paren}{wenyan.answer_span(answer)}{right_paren}"
            return wenyan.answer_span(answer, underline=True, pad_underline=True)

        return blank_pattern.sub(repl, line)

    for line in lines:
        filled_lines.append(fill_line(line))
    return filled_lines


def is_booklet_moxie_noise_line(line: str) -> bool:
    compact = title_plain(line)
    compact = re.sub(r"[_＿\s\u3000]+", "", compact)
    return compact in {"默写必背突围手册", "第一部分分册强化练", "古诗词诵读"}


def normalize_booklet_answer_line(line: str) -> str:
    line = normalize_answer_line(line)
    blank_re = r"(?:[_＿]{2,}|<span\s+class=[\"']blank[\"'][^>]*>\s*</span>)"
    line = re.sub(rf"[（(][^）)]*{blank_re}[^）)]*[）)]", "", line, flags=re.I)
    if "____" not in line and "__" not in line:
        match = re.match(r"^(\s*\d{1,3}[.、]\s*)(.+)$", line)
        if match and re.search(r"(?<=[\u4e00-\u9fff）)])\s+(?=[\u4e00-\u9fff（(])", match.group(2)):
            line = match.group(1) + re.sub(r"\s+", "____", match.group(2).strip())
    return line


BOOKLET_MOXIE_WRAP_UNITS = 92.0
BOOKLET_MOXIE_BREAK_CHARS = set("，。；：？！、,.!?;:")


def booklet_moxie_html_tokens(text: str) -> list[str]:
    tokens: list[str] = []
    pos = 0
    tag_pattern = re.compile(r"<span\b[^>]*>.*?</span>|<sup\b[^>]*>.*?</sup>|<img\b[^>]*>", flags=re.I | re.S)
    for match in tag_pattern.finditer(text):
        if match.start() > pos:
            tokens.extend(text[pos : match.start()])
        tokens.append(match.group(0))
        pos = match.end()
    if pos < len(text):
        tokens.extend(text[pos:])
    return [token for token in tokens if token]


def booklet_moxie_token_units(token: str) -> float:
    visible = wenyan.strip_inline_tags(html.unescape(token))
    if not visible and token.startswith("<img"):
        return 1.0
    return wenyan.text_units(visible)


def booklet_moxie_can_break_after(token: str) -> bool:
    visible = wenyan.strip_inline_tags(html.unescape(token))
    return bool(visible and visible[-1] in BOOKLET_MOXIE_BREAK_CHARS)


def split_booklet_moxie_html_line(text: str, max_units: float = BOOKLET_MOXIE_WRAP_UNITS) -> list[str]:
    if wenyan.text_units(wenyan.strip_inline_tags(text)) <= max_units:
        return [text]
    tokens = booklet_moxie_html_tokens(text)
    fragments: list[str] = []
    current: list[str] = []
    current_units = 0.0
    last_break = -1

    def flush(count: int | None = None) -> None:
        nonlocal current, current_units, last_break
        if count is None:
            count = len(current)
        fragment = "".join(current[:count]).strip()
        if fragment:
            fragments.append(fragment)
        current = current[count:]
        current_units = sum(booklet_moxie_token_units(token) for token in current)
        last_break = -1
        for idx, token in enumerate(current):
            if booklet_moxie_can_break_after(token):
                last_break = idx

    for token in tokens:
        token_units = booklet_moxie_token_units(token)
        if current and current_units + token_units > max_units:
            if last_break >= 0:
                flush(last_break + 1)
            else:
                flush()
        current.append(token)
        current_units += token_units
        if booklet_moxie_can_break_after(token):
            last_break = len(current) - 1
    flush()
    return fragments or [text]


BOOKLET_MOXIE_PAGE_UNITS = 10
BOOKLET_MOXIE_LINE_CAPACITY_UNITS = 42.0


def is_booklet_moxie_subtitle_text(text: str) -> bool:
    compact = title_plain(text)
    return bool(re.match(r"^\d{1,3}\.\s*\S.{0,24}$", compact))


def split_booklet_moxie_styled_lines(
    lines: list[wenyan.StyledLine],
    *,
    moxie_roles: bool = False,
) -> list[wenyan.StyledLine]:
    result: list[wenyan.StyledLine] = []
    for line in lines:
        if line.role == "booklet_moxie_title":
            result.append(wenyan.StyledLine("默写挖空练", line.role, dict(line.meta or {})))
            continue
        role = line.role
        if moxie_roles:
            if is_booklet_moxie_subtitle_text(line.text):
                role = "booklet_moxie_subtitle"
            elif role in {"body", "reading", "extended_reading"}:
                role = "booklet_moxie_body"
        result.append(wenyan.StyledLine(line.text, role, dict(line.meta or {})))
    return result


def booklet_moxie_line_page_cost(line: wenyan.StyledLine) -> int:
    role = line.role or "body"
    if role in {"booklet_moxie_title", "booklet_moxie_subtitle"}:
        return 1
    text = re.sub(r"<img\b[^>]*>", "\u3000", line.text, flags=re.I)
    units = wenyan.text_units(wenyan.strip_inline_tags(html.unescape(text)))
    visual_lines = math.ceil(units / BOOKLET_MOXIE_LINE_CAPACITY_UNITS)
    return max(1, visual_lines)


def chunk_booklet_moxie_lines(lines: list[wenyan.StyledLine], max_units: int = BOOKLET_MOXIE_PAGE_UNITS) -> list[list[wenyan.StyledLine]]:
    header: list[wenyan.StyledLine] = []
    source = lines
    if source and source[0].role == "booklet_moxie_title":
        header = [source[0]]
        source = source[1:]

    groups: list[list[wenyan.StyledLine]] = []
    current: list[wenyan.StyledLine] = []
    for line in source:
        if line.role == "booklet_moxie_subtitle" and current:
            groups.append(current)
            current = [line]
        else:
            current.append(line)
    if current:
        groups.append(current)

    chunks: list[list[wenyan.StyledLine]] = []
    for group_index, group in enumerate(groups):
        pending_header = header if group_index == 0 else []
        header_cost = sum(booklet_moxie_line_page_cost(item) for item in pending_header)
        current_chunk: list[wenyan.StyledLine] = []
        used = 0
        for line in group:
            cost = booklet_moxie_line_page_cost(line)
            limit = max(6, max_units - header_cost)
            if current_chunk and used + cost > limit:
                chunks.append([*pending_header, *current_chunk])
                pending_header = []
                header_cost = 0
                current_chunk = []
                used = 0
                limit = max_units
            current_chunk.append(line)
            used += cost
        if current_chunk or pending_header:
            chunks.append([*pending_header, *current_chunk])
    return chunks


def split_booklet_moxie_piece_groups(lines: list[str]) -> list[tuple[int, str, list[str], str]]:
    starts: list[tuple[int, int, str]] = []
    for index, line in enumerate(lines):
        value = plain(line)
        if not re.fullmatch(r"\d{1,3}", value):
            continue
        lookahead = [title_plain(item) for item in nonempty_after(lines, index, 4)]
        if len(lookahead) >= 2 and not lookahead[0].startswith("教材"):
            starts.append((index, int(value), lookahead[0]))

    groups: list[tuple[int, str, list[str], str]] = []
    for pos, (start, num, title) in enumerate(starts):
        end = starts[pos + 1][0] if pos + 1 < len(starts) else len(lines)
        block = [line for line in lines[start:end] if line.strip()]
        if len(block) < 2:
            continue
        content_start = 2
        if content_start < len(block) and not title_plain(block[content_start]).startswith("教材"):
            content_start += 1
        if content_start < len(block) and title_plain(block[content_start]).startswith("教材"):
            content_start += 1
        display_title = f"{num}.{title_plain(title).replace(' ', '')}"
        content_lines = [
            normalize_body_numbering(line)
            for line in block[content_start:]
            if marker_plain(line) != "默写必背突围手册" and not is_booklet_moxie_noise_line(line)
        ]
        groups.append((num, title, [display_title, *content_lines], canonical_booklet_title(title)))
    return groups


def build_booklet_moxie_body_lines(raw_lines: list[str], poem_index: dict[str, str]) -> list[str]:
    result = ["一、默写挖空练"]
    for _num, _title, display_lines, title_key in split_booklet_moxie_piece_groups(raw_lines):
        answers = derive_moxie_answers_from_poem(display_lines[1:], poem_index.get(title_key, ""))
        result.extend(fill_booklet_moxie_lines(display_lines, answers))
    return result


def booklet_content_summary(block: list[str]) -> str:
    for line in block:
        text = title_plain(line)
        if text.startswith("内含："):
            return text
    return ""


def infer_bibei_tif_search_dirs(*run_dirs: Path) -> list[Path]:
    dirs: list[Path] = []
    for run_dir in run_dirs:
        if not run_dir:
            continue
        dirs.extend(wenyan.infer_tif_search_dirs(run_dir))
        manifest_path = run_dir / "conversion_manifest.json"
        if manifest_path.exists():
            try:
                manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
                for source_dir in manifest.get("tif_source_dirs", []):
                    dirs.append(Path(source_dir))
            except Exception:
                pass
    result: list[Path] = []
    seen: set[str] = set()
    for directory in dirs:
        key = str(directory).lower()
        if key not in seen and directory.exists():
            result.append(directory)
            seen.add(key)
    return result


def open_topic_marker(line: str) -> tuple[int, str] | None:
    value = title_plain(line)
    match = re.match(r"^专题([一二三四五六七八])\s*(.+?)\s*$", value)
    if not match:
        return None
    num = OPEN_TOPIC_NUMERAL_MAP[match.group(1)]
    label = re.sub(r"\s*\d+$", "", match.group(2)).strip()
    return num, f"专题{num}" + (f" {label}" if label else "")


def is_open_figure_placeholder_line(line: str) -> bool:
    value = re.sub(r"\s+", "", plain(line))
    return bool(value and re.fullmatch(r"(?:图\d+)+", value))


def line_has_open_topic_image(line: str) -> bool:
    return bool(re.search(r"<img\b[^>]*(?:data-fbd-tif|src)=", line, flags=re.I))


def is_open_topic_noise_line(line: str) -> bool:
    if line_has_open_topic_image(line):
        return False
    value = marker_plain(line)
    if value in {"默写必背突围手册", "第二部分开放性默写专题练"}:
        return True
    if not plain(line) or plain(line).isdigit():
        return True
    return is_open_figure_placeholder_line(line)


def split_open_topic_pieces(lines: list[str], *, is_answer: bool) -> list[wenyan.Piece]:
    pieces: list[wenyan.Piece] = []
    current_num: int | None = None
    current_title = ""
    current_lines: list[str] = []

    def merge_answer_continuations(content: list[str]) -> list[str]:
        merged: list[str] = []
        for line in content:
            if wenyan.answer_item_match(line) or not merged:
                merged.append(line)
                continue
            if wenyan.answer_item_match(merged[-1]):
                merged[-1] = f"{merged[-1].rstrip()} {line.strip()}"
            else:
                merged.append(line)
        return merged

    def flush() -> None:
        if current_num is None or current_num > OPEN_TOPIC_MAX_NUM:
            return
        normalizer = normalize_answer_line if is_answer else normalize_body_numbering
        content: list[str] = []
        previous_kept_image = False
        for line in current_lines:
            if line_has_open_topic_image(line):
                content.append(normalizer(line))
                previous_kept_image = True
                continue
            if is_open_figure_placeholder_line(line):
                if previous_kept_image:
                    content.append(normalizer(line))
                previous_kept_image = False
                continue
            previous_kept_image = False
            if not is_open_topic_noise_line(line):
                content.append(normalizer(line))
        if is_answer:
            content = merge_answer_continuations(content)
        pieces.append(
            wenyan.Piece(
                volume="册子",
                num=current_num,
                title=current_title,
                author="",
                lines=[current_title, "一、开放性默写专题练", *content],
            )
        )

    for line in lines:
        marker = open_topic_marker(line)
        if marker:
            flush()
            current_num, current_title = marker
            current_lines = []
            if current_num > OPEN_TOPIC_MAX_NUM:
                break
            continue
        if current_num is None:
            continue
        current_lines.append(line)
    flush()
    return pieces


def collect_all_tif_keep_map(source_dir: Path) -> dict[str, str]:
    keep_map: dict[str, str] = {}
    for path in sorted(source_dir.rglob("*.tif")) + sorted(source_dir.rglob("*.tiff")):
        if path.is_file():
            keep_map[normalize_tif_key(path.name)] = tif_markdown_tag(path.name)
    return keep_map


def choose_open_topic_body_fbd(source_dir: Path) -> tuple[Path, str, object] | None:
    if not source_dir.exists():
        return None
    tif_keep_map = collect_all_tif_keep_map(source_dir)
    candidates = [
        source_dir / "必背古诗文册子.fbd",
        source_dir / "必背古诗文.fbd",
    ]
    for candidate in candidates:
        if not candidate.exists():
            continue
        raw = read_fbd(candidate, "gb18030")
        markdown, stats = convert_fbd_to_markdown(raw, tif_replacements=tif_keep_map)
        if "开放性默写" in markdown and "专题一" in markdown:
            return candidate, markdown, stats
    return None


def convert_booklet_open_source_md(
    source_dir: Path = DEFAULT_BOOKLET_OPEN_SOURCE_DIR,
    output_run_dir: Path | None = None,
) -> tuple[Path, Path, list[Path]] | None:
    chosen = choose_open_topic_body_fbd(source_dir)
    answer_fbd = source_dir / "必背古诗文答案.fbd"
    if chosen is None or not answer_fbd.exists():
        return None

    body_fbd, body_md, body_stats = chosen
    tif_keep_map = collect_all_tif_keep_map(source_dir)
    answer_raw = read_fbd(answer_fbd, "gb18030")
    answer_md, answer_stats = convert_fbd_to_markdown(answer_raw, tif_replacements=tif_keep_map)

    run_dir = output_run_dir or DEFAULT_BOOKLET_OPEN_IMAGE_MD_RUN
    run_dir.mkdir(parents=True, exist_ok=True)
    body_path = run_dir / f"正文_{body_fbd.stem}_格式保留.md"
    answer_path = run_dir / f"答案_{answer_fbd.stem}_格式保留.md"
    body_path.write_text(body_md, encoding="utf-8", newline="\n")
    answer_path.write_text(answer_md, encoding="utf-8", newline="\n")
    manifest = {
        "source_dir": str(source_dir),
        "body_fbd": str(body_fbd),
        "answer_fbd": str(answer_fbd),
        "body_md": str(body_path),
        "answer_md": str(answer_path),
        "body_stats": {
            "tif_kept": getattr(body_stats, "tif_kept", 0),
            "tif_unknown": getattr(body_stats, "tif_unknown", 0),
        },
        "answer_stats": {
            "tif_kept": getattr(answer_stats, "tif_kept", 0),
            "tif_unknown": getattr(answer_stats, "tif_unknown", 0),
        },
        "tif_source_dirs": [str(source_dir)],
    }
    (run_dir / "conversion_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return body_path, answer_path, [source_dir]


def open_booklet_md_files_with_images(
    booklet_md_run: Path,
    *,
    source_dir: Path | None = None,
    output_run_dir: Path | None = None,
) -> tuple[Path, Path, list[Path]]:
    converted = convert_booklet_open_source_md(source_dir or DEFAULT_BOOKLET_OPEN_SOURCE_DIR, output_run_dir)
    if converted:
        return converted
    body_md, answer_md = find_md_files(booklet_md_run)
    return body_md, answer_md, []


def parse_booklet_volume_pieces(
    body_md: Path,
    answer_md: Path,
    pocket_md: Path | None = None,
) -> tuple[list[tuple[str, wenyan.Piece, wenyan.Piece | None]], list[wenyan.Piece], list[wenyan.Piece]]:
    body_lines = [line for line in wenyan.clean_lines(body_md.read_text(encoding="utf-8")) if line.strip()]
    answer_lines = [line for line in wenyan.clean_lines(answer_md.read_text(encoding="utf-8")) if line.strip()]
    poem_index = parse_pocket_poem_index(pocket_md)

    body_first = find_marker_index(body_lines, "第一部分分册强化练")
    body_start = find_marker_index(body_lines, "第一部分分册强化练", start=body_first + 1) if body_first >= 0 else -1
    body_second = find_marker_index(body_lines, "第二部分开放性默写专题练", start=body_start + 1 if body_start >= 0 else 0)

    answer_first = find_marker_index(answer_lines, "第一部分分册强化练")
    answer_second = find_marker_index(answer_lines, "第二部分开放性默写专题练", start=answer_first + 1 if answer_first >= 0 else 0)

    volume_entries: list[tuple[str, wenyan.Piece, wenyan.Piece | None]] = []
    for idx, (volume_label, _filename) in enumerate(BIBEI_BOOKLET_VOLUME_LABELS, 1):
        next_label = BIBEI_BOOKLET_VOLUME_LABELS[idx][0] if idx < len(BIBEI_BOOKLET_VOLUME_LABELS) else None
        b_start = find_marker_index(body_lines, volume_label, start=body_start)
        b_end = find_marker_index(body_lines, next_label, start=b_start + 1) if next_label else body_second
        if b_start < 0 or b_end < 0:
            continue
        block = body_lines[b_start + 1 : b_end]
        m_idx = next((i for i, line in enumerate(block) if marker_plain(line) == "默写挖空练"), -1)
        k_idx = next((i for i, line in enumerate(block) if marker_plain(line) == "重点语句强化练"), -1)
        if m_idx < 0 or k_idx < 0:
            continue
        summary = booklet_content_summary(block[:m_idx])
        moxie_lines = build_booklet_moxie_body_lines(block[m_idx + 1 : k_idx], poem_index)
        body_piece = wenyan.Piece(
            volume="册子",
            num=idx,
            title=f"分册强化练{volume_label}",
            author=summary,
            lines=[
                f"分册强化练{volume_label}",
                *moxie_lines,
                "二、重点语句强化练",
                *[normalize_body_numbering(line) for line in block[k_idx + 1 :] if not plain(line).isdigit()],
            ],
        )

        answer_piece: wenyan.Piece | None = None
        a_start = find_marker_index(answer_lines, volume_label, start=answer_first)
        a_end = find_marker_index(answer_lines, next_label, start=a_start + 1) if next_label and a_start >= 0 else answer_second
        if a_start >= 0 and a_end > a_start:
            a_block = answer_lines[a_start + 1 : a_end]
            a_key = next((i for i, line in enumerate(a_block) if marker_plain(line) == "重点语句强化练"), -1)
            if a_key >= 0:
                answer_piece = wenyan.Piece(
                    volume="册子",
                    num=idx,
                    title=f"分册强化练{volume_label}",
                    author="",
                    lines=[
                        f"分册强化练{volume_label}",
                        "二、重点语句强化练",
                        *[normalize_booklet_answer_line(line) for line in a_block[a_key + 1 :]],
                    ],
                )
        volume_entries.append((volume_label, body_piece, answer_piece))

    open_body_pieces: list[wenyan.Piece] = []
    open_answer_pieces: list[wenyan.Piece] = []
    if body_second >= 0:
        open_lines = body_lines[body_second + 1 :]
        open_body_pieces = split_open_topic_pieces(open_lines, is_answer=False)
    if answer_second >= 0:
        open_ans = answer_lines[answer_second + 1 :]
        open_answer_pieces = split_open_topic_pieces(open_ans, is_answer=True)
    return volume_entries, open_body_pieces, open_answer_pieces


def scan_bibei_main_group_units(main_md_run: Path) -> list[dict[str, object]]:
    """Scan all independently selectable main-book units from converted MD."""
    body_md, _answer_md = find_md_files(main_md_run)
    pieces = parse_body_pieces(body_md)
    extra_pieces = parse_main_extra_body_pieces(body_md)
    pieces.extend(extra_pieces)
    rush_piece = parse_main_rush_piece(body_md)
    if rush_piece:
        pieces.append(rush_piece)
    return [
        {"volume": piece.volume, "num": piece.num, "title": piece.title}
        for piece in pieces
    ]


def scan_bibei_booklet_group_units(
    booklet_md_run: Path,
    *,
    booklet_open_source_dir: Path | None = None,
    booklet_open_image_md_run: Path | None = None,
) -> list[dict[str, object]]:
    """Scan booklet divisions and open-topic units from their respective body MD."""
    body_md, answer_md = find_md_files(booklet_md_run)
    entries, _open_body, _open_answer = parse_booklet_volume_pieces(body_md, answer_md)
    open_body_md, open_answer_md, _extra_tif_dirs = open_booklet_md_files_with_images(
        booklet_md_run,
        source_dir=booklet_open_source_dir,
        output_run_dir=booklet_open_image_md_run,
    )
    _entries, open_body, _open_answers = parse_booklet_volume_pieces(open_body_md, open_answer_md)
    units = [
        {"volume": f"分册：{volume_label}", "num": piece.num, "title": piece.title}
        for volume_label, piece, _answer in entries
    ]
    units.extend(
        {"volume": "开放性默写专题练", "num": piece.num, "title": piece.title}
        for piece in open_body
    )
    return units


def generate_bibei_main_complete_pptx(
    *,
    main_md_run: Path,
    template: Path,
    output_root: Path,
    group_ranges: list[dict[str, object]] | None = None,
    progress: Callable[[str], None] | None = None,
) -> list[Path]:
    def emit(message: str) -> None:
        if progress:
            progress(message)
        else:
            print(message)

    main_body_md, main_answer_md = find_md_files(main_md_run)
    tif_search_dirs = wenyan.infer_tif_search_dirs(main_md_run)
    main_dir = output_root / "主书"
    main_dir.mkdir(parents=True, exist_ok=True)

    if group_ranges is not None:
        body_pieces = parse_body_pieces(main_body_md)
        answer_pieces = parse_answer_pieces(main_answer_md)
        answer_map = build_bibei_answer_piece_map(body_pieces, answer_pieces)
        extra_body = parse_main_extra_body_pieces(main_body_md)
        extra_answer = parse_main_extra_answer_pieces(main_answer_md, extra_body)
        answer_map.update({wenyan.piece_key(piece): piece for piece in extra_answer})
        rush_body = parse_main_rush_piece(main_body_md)
        rush_answer = parse_main_rush_answer_piece(main_answer_md)
        all_pieces = [*body_pieces, *extra_body]
        if rush_body:
            all_pieces.append(rush_body)
            if rush_answer:
                answer_map[wenyan.piece_key(rush_body)] = rush_answer
        by_key = {wenyan.piece_key(piece): piece for piece in all_pieces}
        outputs: list[Path] = []
        for group_index, group in enumerate(group_ranges, 1):
            keys = [tuple(key) for key in group["unit_keys"]]
            selected = [by_key[key] for key in keys if key in by_key]
            if len(selected) != len(keys):
                missing = [key for key in keys if key not in by_key]
                raise ValueError(f"主书分组中存在无法回查的 MD 篇目：{missing}")
            first, last = selected[0], selected[-1]
            is_core_group = all(1 <= piece.num <= 60 for piece in selected)
            suffix = "+「6+1训练赢」" if is_core_group else ""
            filename = f"{group_index} 篇目{first.num}-{last.num}{suffix}.pptx"
            output = main_dir / filename
            emit(f"生成：{filename}")
            build_bibei_deck(
                template,
                output,
                selected,
                answer_map,
                volume=first.volume,
                start=first.num,
                end=last.num,
                tif_search_dirs=tif_search_dirs,
                include_training_label=is_core_group,
                cover_subtitle=None if is_core_group else f"篇目{first.num}-{last.num}",
            )
            outputs.append(output)
        manifest = {
            "main_md_run": str(main_md_run),
            "template": str(template),
            "output_root": str(output_root),
            "group_ranges": group_ranges,
            "outputs": [str(path) for path in outputs],
        }
        (output_root / "bibei_main_complete_generation_manifest.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return outputs

    outputs, _stats = generate_bibei_main_pptx_all(md_run=main_md_run, template=template, output_dir=main_dir, progress=emit)

    extra_body = parse_main_extra_body_pieces(main_body_md)
    extra_answer = parse_main_extra_answer_pieces(main_answer_md, extra_body)
    if extra_body:
        emit("生成：11 篇目61-72.pptx")
        output = main_dir / "11 篇目61-72.pptx"
        build_bibei_deck(
            template,
            output,
            extra_body,
            {wenyan.piece_key(piece): piece for piece in extra_answer},
            volume="附赠选修12篇",
            start=61,
            end=72,
            tif_search_dirs=tif_search_dirs,
            include_training_label=False,
            cover_subtitle="篇目61-72",
        )
        outputs.append(output)

    rush_body = parse_main_rush_piece(main_body_md)
    rush_answer = parse_main_rush_answer_piece(main_answer_md)
    if rush_body and rush_answer:
        emit("生成：12 150组冲刺练.pptx")
        output = main_dir / "12 150组冲刺练.pptx"
        build_bibei_deck(
            template,
            output,
            [rush_body],
            {wenyan.piece_key(rush_answer): rush_answer},
            volume="主书",
            start=1,
            end=150,
            tif_search_dirs=tif_search_dirs,
            include_training_label=False,
            cover_subtitle="150组冲刺练",
        )
        outputs.append(output)

    manifest = {
        "main_md_run": str(main_md_run),
        "template": str(template),
        "output_root": str(output_root),
        "outputs": [str(path) for path in outputs],
    }
    output_root.mkdir(parents=True, exist_ok=True)
    (output_root / "bibei_main_complete_generation_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return outputs


def generate_bibei_complete_pptx(
    *,
    main_md_run: Path,
    booklet_md_run: Path,
    template: Path,
    output_root: Path,
    booklet_open_source_dir: Path | None = None,
    booklet_open_image_md_run: Path | None = None,
    progress: Callable[[str], None] | None = None,
) -> list[Path]:
    def emit(message: str) -> None:
        if progress:
            progress(message)
        else:
            print(message)

    main_body_md, main_answer_md = find_md_files(main_md_run)
    booklet_body_md, booklet_answer_md = find_md_files(booklet_md_run)
    tif_search_dirs = list(dict.fromkeys([*wenyan.infer_tif_search_dirs(main_md_run), *wenyan.infer_tif_search_dirs(booklet_md_run)]))

    main_dir = output_root / "主书"
    booklet_dir = output_root / "册子"
    main_dir.mkdir(parents=True, exist_ok=True)
    booklet_dir.mkdir(parents=True, exist_ok=True)

    outputs = generate_bibei_main_complete_pptx(
        main_md_run=main_md_run,
        template=template,
        output_root=output_root,
        progress=emit,
    )

    booklet_entries, open_body_pieces, open_answer_pieces = parse_booklet_volume_pieces(booklet_body_md, booklet_answer_md)
    for _volume_label, body_piece, answer_piece in booklet_entries:
        filename = BIBEI_BOOKLET_VOLUME_LABELS[body_piece.num - 1][1]
        emit(f"生成：{filename}")
        output = booklet_dir / filename
        answer_map = {wenyan.piece_key(answer_piece): answer_piece} if answer_piece else {}
        build_bibei_deck(
            template,
            output,
            [body_piece],
            answer_map,
            volume="册子",
            start=body_piece.num,
            end=body_piece.num,
            tif_search_dirs=tif_search_dirs,
            include_training_label=False,
            cover_subtitle=body_piece.title,
        )
        outputs.append(output)
    open_body_md_with_images, open_answer_md_with_images, open_extra_tif_dirs = open_booklet_md_files_with_images(
        booklet_md_run,
        source_dir=booklet_open_source_dir,
        output_run_dir=booklet_open_image_md_run,
    )
    if open_extra_tif_dirs:
        tif_search_dirs = list(dict.fromkeys([*tif_search_dirs, *open_extra_tif_dirs]))
    _unused_entries, open_body_pieces, open_answer_pieces = parse_booklet_volume_pieces(open_body_md_with_images, open_answer_md_with_images)
    if open_body_pieces:
        emit("生成：7 开放性默写专题练专题一到八.pptx")
        output = booklet_dir / "7 开放性默写专题练专题一到八.pptx"
        answer_map = {wenyan.piece_key(piece): piece for piece in open_answer_pieces}
        build_bibei_deck(
            template,
            output,
            open_body_pieces,
            answer_map,
            volume="册子",
            start=1,
            end=8,
            tif_search_dirs=tif_search_dirs,
            include_training_label=False,
            cover_subtitle="开放性默写专题练 专题1-8",
            directory_page_size=8,
        )
        outputs.append(output)

    manifest = {
        "main_md_run": str(main_md_run),
        "booklet_md_run": str(booklet_md_run),
        "template": str(template),
        "output_root": str(output_root),
        "outputs": [str(path) for path in outputs],
    }
    output_root.mkdir(parents=True, exist_ok=True)
    (output_root / "bibei_complete_generation_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return outputs


def generate_bibei_booklet_pptx(
    *,
    booklet_md_run: Path,
    pocket_md_run: Path,
    template: Path,
    output_dir: Path,
    group_ranges: list[dict[str, object]] | None = None,
    limit_volumes: int | None = None,
    booklet_open_source_dir: Path | None = None,
    booklet_open_image_md_run: Path | None = None,
    progress: Callable[[str], None] | None = None,
) -> list[Path]:
    def emit(message: str) -> None:
        if progress:
            progress(message)
        else:
            print(message)

    booklet_body_md, booklet_answer_md = find_md_files(booklet_md_run)
    pocket_md = find_pocket_md_file(pocket_md_run)
    tif_search_dirs = infer_bibei_tif_search_dirs(booklet_md_run, pocket_md_run)
    output_dir.mkdir(parents=True, exist_ok=True)

    emit(f"册子正文 MD：{booklet_body_md}")
    emit(f"册子答案 MD：{booklet_answer_md}")
    emit(f"口袋本 MD：{pocket_md}")
    emit(f"模板 PPTX：{template}")

    booklet_entries, open_body_pieces, open_answer_pieces = parse_booklet_volume_pieces(booklet_body_md, booklet_answer_md, pocket_md)
    outputs: list[Path] = []

    if group_ranges is not None:
        open_body_md, open_answer_md, open_extra_tif_dirs = open_booklet_md_files_with_images(
            booklet_md_run,
            source_dir=booklet_open_source_dir,
            output_run_dir=booklet_open_image_md_run,
        )
        if open_extra_tif_dirs:
            tif_search_dirs = list(dict.fromkeys([*tif_search_dirs, *open_extra_tif_dirs]))
        _unused_entries, open_body_pieces, open_answer_pieces = parse_booklet_volume_pieces(
            open_body_md, open_answer_md, pocket_md
        )
        units_by_key: dict[tuple[str, int], tuple[wenyan.Piece, wenyan.Piece | None]] = {}
        for volume_label, body_piece, answer_piece in booklet_entries:
            units_by_key[(f"分册：{volume_label}", body_piece.num)] = (body_piece, answer_piece)
        open_answer_by_num = {piece.num: piece for piece in open_answer_pieces}
        for body_piece in open_body_pieces:
            units_by_key[("开放性默写专题练", body_piece.num)] = (body_piece, open_answer_by_num.get(body_piece.num))

        for group_index, group in enumerate(group_ranges, 1):
            keys = [tuple(key) for key in group["unit_keys"]]
            selected_pairs = [units_by_key[key] for key in keys if key in units_by_key]
            if len(selected_pairs) != len(keys):
                missing = [key for key in keys if key not in units_by_key]
                raise ValueError(f"小册子分组中存在无法回查的 MD 篇目：{missing}")
            selected_pieces = [pair[0] for pair in selected_pairs]
            answer_map = {
                wenyan.piece_key(answer_piece): answer_piece
                for _body_piece, answer_piece in selected_pairs
                if answer_piece is not None
            }
            first, last = selected_pieces[0], selected_pieces[-1]
            if first.title.startswith("分册强化练"):
                display_range = first.title.replace("分册强化练", "")
            elif len(selected_pieces) == 1:
                display_range = first.title
            else:
                display_range = f"开放性默写专题练 专题{first.num}-{last.num}"
            if first.title.startswith("分册强化练") and len(selected_pieces) == 1:
                filename = f"{group_index} {first.title}.pptx"
            elif (
                first.title.startswith("专题")
                and first.num == 1
                and last.num == 8
                and len(selected_pieces) == 8
            ):
                filename = f"{group_index} 开放性默写专题练专题一到八.pptx"
            else:
                filename = f"{group_index} {display_range}.pptx"
            output = output_dir / filename
            emit(f"生成：{filename}")
            build_bibei_deck(
                template,
                output,
                selected_pieces,
                answer_map,
                volume="册子",
                start=first.num,
                end=last.num,
                tif_search_dirs=tif_search_dirs,
                include_training_label=False,
                cover_subtitle=display_range,
                directory_page_size=max(6, len(selected_pieces)),
            )
            outputs.append(output)

        manifest = {
            "booklet_md_run": str(booklet_md_run),
            "pocket_md_run": str(pocket_md_run),
            "template": str(template),
            "output_dir": str(output_dir),
            "group_ranges": group_ranges,
            "outputs": [str(path) for path in outputs],
        }
        (output_dir / "bibei_booklet_generation_manifest.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return outputs

    selected_entries = booklet_entries if limit_volumes is None else booklet_entries[:limit_volumes]
    for _volume_label, body_piece, answer_piece in selected_entries:
        filename = BIBEI_BOOKLET_VOLUME_LABELS[body_piece.num - 1][1]
        output = output_dir / filename
        answer_map = {wenyan.piece_key(answer_piece): answer_piece} if answer_piece else {}
        emit(f"生成：{filename}")
        build_bibei_deck(
            template,
            output,
            [body_piece],
            answer_map,
            volume="册子",
            start=body_piece.num,
            end=body_piece.num,
            tif_search_dirs=tif_search_dirs,
            include_training_label=False,
            cover_subtitle=body_piece.title,
        )
        outputs.append(output)

    if limit_volumes is None:
        open_body_md_with_images, open_answer_md_with_images, open_extra_tif_dirs = open_booklet_md_files_with_images(
            booklet_md_run,
            source_dir=booklet_open_source_dir,
            output_run_dir=booklet_open_image_md_run,
        )
        if open_extra_tif_dirs:
            tif_search_dirs = list(dict.fromkeys([*tif_search_dirs, *open_extra_tif_dirs]))
        _unused_entries, open_body_pieces, open_answer_pieces = parse_booklet_volume_pieces(
            open_body_md_with_images,
            open_answer_md_with_images,
            pocket_md,
        )

    if limit_volumes is None and open_body_pieces:
        output = output_dir / "7 开放性默写专题练专题一到八.pptx"
        answer_map = {wenyan.piece_key(piece): piece for piece in open_answer_pieces}
        emit("生成：7 开放性默写专题练专题一到八.pptx")
        build_bibei_deck(
            template,
            output,
            open_body_pieces,
            answer_map,
            volume="册子",
            start=1,
            end=8,
            tif_search_dirs=tif_search_dirs,
            include_training_label=False,
            cover_subtitle="开放性默写专题练 专题1-8",
            directory_page_size=8,
        )
        outputs.append(output)

    manifest = {
        "booklet_md_run": str(booklet_md_run),
        "pocket_md_run": str(pocket_md_run),
        "template": str(template),
        "output_dir": str(output_dir),
        "limit_volumes": limit_volumes,
        "outputs": [str(path) for path in outputs],
    }
    (output_dir / "bibei_booklet_generation_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return outputs


def generate_bibei_booklet_open_pptx(
    *,
    booklet_md_run: Path,
    pocket_md_run: Path,
    template: Path,
    output_dir: Path,
    booklet_open_source_dir: Path | None = None,
    booklet_open_image_md_run: Path | None = None,
    progress: Callable[[str], None] | None = None,
) -> Path:
    def emit(message: str) -> None:
        if progress:
            progress(message)
        else:
            print(message)

    booklet_body_md, booklet_answer_md, open_extra_tif_dirs = open_booklet_md_files_with_images(
        booklet_md_run,
        source_dir=booklet_open_source_dir,
        output_run_dir=booklet_open_image_md_run,
    )
    pocket_md = find_pocket_md_file(pocket_md_run)
    tif_search_dirs = infer_bibei_tif_search_dirs(booklet_md_run, pocket_md_run)
    if open_extra_tif_dirs:
        tif_search_dirs = list(dict.fromkeys([*tif_search_dirs, *open_extra_tif_dirs]))
    _booklet_entries, open_body_pieces, open_answer_pieces = parse_booklet_volume_pieces(
        booklet_body_md,
        booklet_answer_md,
        pocket_md,
    )
    if not open_body_pieces:
        raise ValueError("未从小册子正文 MD 中解析到开放性默写专题练")

    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / "7 开放性默写专题练专题一到八.pptx"
    emit(f"开放性专题数：{len(open_body_pieces)}")
    emit(f"开放性答案专题数：{len(open_answer_pieces)}")
    emit(f"生成：{output.name}")
    build_bibei_deck(
        template,
        output,
        open_body_pieces,
        {wenyan.piece_key(piece): piece for piece in open_answer_pieces},
        volume="册子",
        start=1,
        end=8,
        tif_search_dirs=tif_search_dirs,
        include_training_label=False,
        cover_subtitle="开放性默写专题练 专题1-8",
        directory_page_size=8,
    )
    return output


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="测试：必背古诗文 md -> PPTX")
    parser.add_argument("--md-run", default=str(DEFAULT_MD_RUN), help="必背古诗文 md run 目录")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="必背古诗文 PPTX 模板")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="测试 PPTX 输出目录")
    parser.add_argument("--limit-pieces", type=int, default=6, help="生成前几个篇目，默认 6")
    parser.add_argument("--all-main", action="store_true", help="生成篇目1-60的10个正式PPTX")
    parser.add_argument("--main-complete", action="store_true", help="只生成主书：篇目1-72与150组冲刺练")
    parser.add_argument("--complete", action="store_true", help="按主书/册子完整结构生成全部PPTX")
    parser.add_argument("--booklet-md-run", default="", help="必背古诗文册子 md run 目录")
    parser.add_argument("--pocket-md-run", default=str(DEFAULT_POCKET_MD_RUN), help="必背口袋本 md run 目录")
    parser.add_argument("--booklet-only", action="store_true", help="只生成小册子 PPTX")
    parser.add_argument("--booklet-open-only", action="store_true", help="只生成小册子的开放性默写专题练 PPTX")
    parser.add_argument("--booklet-limit", type=int, default=0, help="小册子生成前几个分册；0 表示全部")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.all_main:
        outputs, stats = generate_bibei_main_pptx_all(
            md_run=Path(args.md_run),
            template=Path(args.template),
            output_dir=Path(args.output_dir),
        )
        print(f"已生成 {len(outputs)} 个 PPTX，目录：{Path(args.output_dir)}")
        print(f"解析结果：正文 {stats.body_count} 篇，答案 {stats.answer_count} 篇")
        return 0
    if args.main_complete:
        outputs = generate_bibei_main_complete_pptx(
            main_md_run=Path(args.md_run),
            template=Path(args.template),
            output_root=Path(args.output_dir),
        )
        print(f"已生成 {len(outputs)} 个主书 PPTX，目录：{Path(args.output_dir) / '主书'}")
        return 0
    if args.complete:
        if not args.booklet_md_run:
            raise ValueError("--complete 需要提供 --booklet-md-run")
        outputs = generate_bibei_complete_pptx(
            main_md_run=Path(args.md_run),
            booklet_md_run=Path(args.booklet_md_run),
            template=Path(args.template),
            output_root=Path(args.output_dir),
        )
        print(f"已生成 {len(outputs)} 个 PPTX，目录：{Path(args.output_dir)}")
        return 0
    if args.booklet_only:
        outputs = generate_bibei_booklet_pptx(
            booklet_md_run=Path(args.booklet_md_run or DEFAULT_BOOKLET_MD_RUN),
            pocket_md_run=Path(args.pocket_md_run),
            template=Path(args.template),
            output_dir=Path(args.output_dir),
            limit_volumes=args.booklet_limit or None,
        )
        print(f"已生成 {len(outputs)} 个小册子 PPTX，目录：{Path(args.output_dir)}")
        return 0
    if args.booklet_open_only:
        output = generate_bibei_booklet_open_pptx(
            booklet_md_run=Path(args.booklet_md_run or DEFAULT_BOOKLET_MD_RUN),
            pocket_md_run=Path(args.pocket_md_run),
            template=Path(args.template),
            output_dir=Path(args.output_dir),
        )
        print(f"已生成小册子开放性默写专题练 PPTX：{output}")
        return 0
    output, stats = generate_bibei_test_pptx(
        md_run=Path(args.md_run),
        template=Path(args.template),
        output_dir=Path(args.output_dir),
        limit_pieces=args.limit_pieces,
    )
    print(f"已生成：{output}")
    print(f"解析结果：正文 {stats.body_count} 篇，答案 {stats.answer_count} 篇；本次生成：{', '.join(stats.first_titles)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

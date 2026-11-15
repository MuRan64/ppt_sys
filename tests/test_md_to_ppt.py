#!/usr/bin/env python3
"""Experimental Markdown-to-PPTX generator for 教材文言梳理.

This is a test script only. It is intentionally separate from the GUI/main
program while we learn the content-to-slide rules.

The script:
- reads the converted body/answer Markdown files,
- pairs pieces by volume + piece number + title,
- clones a PowerPoint template at the Open XML level,
- generates the first one or two split decks for visual inspection.
"""

from __future__ import annotations

import argparse
import hashlib
import html
import io
import json
import math
import os
import posixpath
import re
import shutil
import subprocess
import sys
import zipfile
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Callable, Iterable
from xml.etree import ElementTree as ET


PROJECT_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
DEFAULT_MD_RUN = PROJECT_DIR / "out_md" / "run_20260707_172914"
DEFAULT_OUTPUT_DIR = PROJECT_DIR / "output" / "test_ppt"
DEFAULT_TEMPLATE = PROJECT_DIR / "文言梳理pptx模板.pptx"
MODULE_IMAGE_DIR = PROJECT_DIR / "tests" / "static" / "文言梳理"

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_IMAGE_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
REL_SLIDE_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"

ET.register_namespace("p", P_NS)
ET.register_namespace("a", A_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("", CT_NS)

EMU_PER_INCH = 914400
SLIDE_W_IN = 13.333333
SLIDE_H_IN = 7.5

VOLUMES = [
    "必修上册",
    "必修下册",
    "选择性必修上册",
    "选择性必修中册",
    "选择性必修下册",
    "教材外必背篇目",
]

GROUPS = [
    ("1必修上册篇目1-4.pptx", "必修上册", 1, 4),
    ("2必修上册篇目5-8.pptx", "必修上册", 5, 8),
    ("3必修下册篇目1-4.pptx", "必修下册", 1, 4),
    ("4必修下册篇目5-8.pptx", "必修下册", 5, 8),
    ("5必修下册篇目9-11.pptx", "必修下册", 9, 11),
    ("6必修下册篇目12-14.pptx", "必修下册", 12, 14),
    ("7选择性必修上册篇目1-4.pptx", "选择性必修上册", 1, 4),
    ("8选择性必修上册篇目5-8.pptx", "选择性必修上册", 5, 8),
    ("9选择性必修中册篇目1-3.pptx", "选择性必修中册", 1, 3),
    ("10选择性必修中册篇目4-6.pptx", "选择性必修中册", 4, 6),
    ("11选择性必修下册篇目1-3.pptx", "选择性必修下册", 1, 3),
    ("12选择性必修下册篇目4-6.pptx", "选择性必修下册", 4, 6),
    ("13选择性必修下册篇目7-9.pptx", "选择性必修下册", 7, 9),
    ("14选择性必修下册篇目10-13.pptx", "选择性必修下册", 10, 13),
    ("15教材外必背篇目篇目1-4.pptx", "教材外必背篇目", 1, 4),
]


@dataclass
class InlineRun:
    text: str
    underline: str | None = None
    underline_color: str | None = None
    highlight: str | None = None
    color: str | None = None
    font: str | None = None
    bold: bool | None = None
    superscript: bool = False
    answer_fill: bool = False
    emphasis_dot: bool = False
    tif_name: str | None = None


@dataclass
class StyledLine:
    text: str
    role: str = "body"  # heading/subheading/center_heading/reading/body/answer/analysis
    meta: dict[str, object] | None = None


@dataclass
class LineBlock:
    lines: list[StyledLine]
    kind: str = "single"


@dataclass
class AnimationTarget:
    shape_id: int
    effect: str = "appear"  # appear/flyin
    text_target: bool = True


@dataclass
class AnswerOverlaySpec:
    ranges: list[tuple[int, int]]
    text_lines: list[str]
    expected_width_pt: float
    compact: bool = False
    inline_context: bool = False


@dataclass
class Piece:
    volume: str
    num: int
    title: str
    author: str
    lines: list[str]


@dataclass
class AnswerItem:
    answers: list[str]
    analyses: list[str]


class InlineMarkdownParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.runs: list[InlineRun] = []
        self.stack: list[dict[str, object]] = []
        self.current_underline: str | None = None
        self.current_underline_color: str | None = None
        self.current_highlight: str | None = None
        self.current_color: str | None = None
        self.current_font: str | None = None
        self.current_bold: bool | None = None
        self.current_dot = False
        self.current_sup = False
        self.current_answer_fill = False

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attr = dict(attrs)
        if tag == "img":
            self._append_tif_run(attr)
            return
        state = {
            "underline": self.current_underline,
            "underline_color": self.current_underline_color,
            "highlight": getattr(self, "current_highlight", None),
            "color": self.current_color,
            "font": self.current_font,
            "bold": self.current_bold,
            "dot": self.current_dot,
            "sup": self.current_sup,
            "answer_fill": self.current_answer_fill,
        }
        self.stack.append(state)
        if tag == "sup":
            self.current_sup = True
        elif tag == "span":
            classes = set((attr.get("class", "") or "").split())
            if "dot" in classes:
                self.current_dot = True
            if "uline" in classes:
                self.current_underline = "sng"
            if "wave" in classes:
                self.current_underline = "wavy"
                self.current_underline_color = "00AEEF"
            if classes & {"blue-underline", "blue-line", "reading-underline"}:
                self.current_underline_color = "00AEEF"
            if "yuedu-bg-word" in classes:
                self.current_highlight = "D9D9D9"
            if classes & {"blue-text", "blue-term", "sanbu-keyword"}:
                self.current_color = "00AEEF"
            if classes & {"font-song", "song-font"}:
                self.current_font = "宋体"
            if classes & {"font-kaiti", "kaiti-font"}:
                self.current_font = "楷体"
            if classes & {"font-times", "times-font"}:
                self.current_font = "Times New Roman"
            if classes & {"font-heiti", "heiti-font"}:
                self.current_font = "黑体"
            if "bold-text" in classes:
                self.current_bold = True
            if "normal-weight" in classes:
                self.current_bold = False
            if "answer-fill" in classes:
                self.current_answer_fill = True

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag == "img":
            self._append_tif_run(dict(attrs))
            return
        self.handle_starttag(tag, attrs)
        self.handle_endtag(tag)

    def handle_endtag(self, tag: str) -> None:
        if self.stack:
            state = self.stack.pop()
            self.current_underline = state["underline"]  # type: ignore[assignment]
            self.current_underline_color = state["underline_color"]  # type: ignore[assignment]
            self.current_highlight = state["highlight"]  # type: ignore[assignment]
            self.current_color = state["color"]  # type: ignore[assignment]
            self.current_font = state["font"]  # type: ignore[assignment]
            self.current_bold = state["bold"]  # type: ignore[assignment]
            self.current_dot = bool(state["dot"])
            self.current_sup = bool(state["sup"])
            self.current_answer_fill = bool(state["answer_fill"])

    def handle_data(self, data: str) -> None:
        if not data:
            return
        self.runs.append(
            InlineRun(
                text=data,
                underline=self.current_underline,
                underline_color=(
                    "000000"
                    if self.current_answer_fill and self.current_underline
                    else self.current_underline_color
                ),
                highlight=getattr(self, "current_highlight", None),
                color=self.current_color,
                font=self.current_font,
                bold=self.current_bold,
                superscript=self.current_sup,
                answer_fill=self.current_answer_fill,
                emphasis_dot=self.current_dot,
            )
        )

    def _append_tif_run(self, attr: dict[str, str | None]) -> None:
        classes = set((attr.get("class", "") or "").split())
        if "fbd-tif-char" not in classes and not attr.get("data-fbd-tif"):
            return
        tif_name = attr.get("data-fbd-tif") or attr.get("src") or attr.get("alt")
        if not tif_name:
            return
        self.runs.append(
            InlineRun(
                text="\u3000",
                underline=self.current_underline,
                underline_color=self.current_underline_color if self.current_underline else None,
                superscript=self.current_sup,
                answer_fill=False,
                emphasis_dot=self.current_dot,
                tif_name=Path(tif_name).name,
            )
        )


def add_emphasis_dot(text: str) -> str:
    """Legacy text fallback; normal PPT output now uses separate dot runs."""
    chars: list[str] = []
    for ch in text:
        chars.append(ch)
        if re.match(r"[\u4e00-\u9fffA-Za-z0-9]", ch):
            chars.append("．")
    return "".join(chars)


def inch(value: float) -> int:
    return int(round(value * EMU_PER_INCH))


def qn(ns: str, name: str) -> str:
    return f"{{{ns}}}{name}"


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def strip_style_block(md: str) -> str:
    return re.sub(r"(?is)<style.*?</style>", "", md).strip()


def blank_span_to_underscores(match: re.Match[str]) -> str:
    style = match.group(1) or ""
    width_match = re.search(r"width:\s*([0-9.]+)em", style)
    width = float(width_match.group(1)) if width_match else 3.0
    count = max(3, min(18, int(round(width * 1.8))))
    return "_" * count


def normalize_md_line(line: str) -> str:
    line = re.sub(
        r'<span\s+class="blank"([^>]*)></span>',
        blank_span_to_underscores,
        line,
    )
    line = html.unescape(line)
    return line.strip()


def clean_lines(md: str) -> list[str]:
    text = strip_style_block(md)
    return [normalize_md_line(line) for line in text.splitlines()]


def nonempty_after(lines: list[str], index: int, count: int = 4) -> list[str]:
    found: list[str] = []
    for line in lines[index + 1 :]:
        if line:
            found.append(line)
        if len(found) >= count:
            break
    return found


def is_timed_detection_title(title: str, current_volume: str | None) -> bool:
    return bool(current_volume and title == f"{current_volume}定时检测")


def is_piece_start(lines: list[str], index: int, current_volume: str | None, *, is_answer: bool) -> re.Match[str] | None:
    if not current_volume:
        return None
    line = lines[index]
    match = re.match(r"^(\d+)\.\s*(.+?)\s*$", line)
    if not match:
        return None
    title = match.group(2)
    if not title or title.startswith("①"):
        return None
    if is_timed_detection_title(title, current_volume):
        return match
    lookahead = nonempty_after(lines, index, 3)
    # Real piece starts are followed immediately by either:
    #   title -> 课文梳理                 (answer file)
    #   title -> author -> 课文梳理        (body file)
    # A loose "课文梳理 appears somewhere soon" check incorrectly catches
    # numbered questions near the end of the previous piece.
    if is_answer:
        if lookahead and lookahead[0] == "课文梳理":
            return match
        return None
    if lookahead and (
        lookahead[0] == "课文梳理"
        or (len(lookahead) > 1 and lookahead[1] == "课文梳理" and not re.match(r"^\d+\.\s*", lookahead[0]))
    ):
        return match
    return None


def parse_pieces(md_path: Path, is_answer: bool) -> list[Piece]:
    lines = clean_lines(md_path.read_text(encoding="utf-8"))
    starts: list[tuple[int, str, int, str]] = []
    current_volume: str | None = None
    for i, line in enumerate(lines):
        if line in VOLUMES:
            current_volume = line
            continue
        match = is_piece_start(lines, i, current_volume, is_answer=is_answer)
        if match and current_volume:
            title = re.sub(r"_+", "", match.group(2)).strip()
            starts.append((i, current_volume, int(match.group(1)), title))

    pieces: list[Piece] = []
    for idx, (start, volume, num, title) in enumerate(starts):
        end = starts[idx + 1][0] if idx + 1 < len(starts) else len(lines)
        block = [line for line in lines[start:end] if line]
        if not block:
            continue
        author = ""
        if not is_answer and len(block) >= 2 and block[1] != "课文梳理":
            author = block[1]
        pieces.append(Piece(volume=volume, num=num, title=title, author=author, lines=block))
    return pieces


def piece_key(piece: Piece) -> tuple[str, int]:
    return piece.volume, piece.num


def split_sections(piece: Piece) -> list[tuple[str, list[str]]]:
    section_titles = re.compile(r"^[一二三四五六七八九十]+、")
    module_titles = {"课文梳理", "知识梳理", "情境默写", "拓展阅读"}
    sections: list[tuple[str, list[str]]] = []
    current_title = "导入"
    current_lines: list[str] = []

    # Skip piece title, author, and keep from module marker onward.
    source = piece.lines[1:]
    if piece.author and source and source[0] == piece.author:
        source = source[1:]

    for line in source:
        if line in module_titles or section_titles.match(line):
            if current_lines:
                sections.append((current_title, current_lines))
            current_title = line
            current_lines = [line]
        else:
            current_lines.append(line)
    if current_lines:
        sections.append((current_title, current_lines))
    return sections


def answer_section_map(answer_piece: Piece | None) -> dict[str, list[str]]:
    if not answer_piece:
        return {}
    sections = split_sections_with_context(answer_piece)
    result: dict[str, list[str]] = {}
    for module, title, lines in sections:
        result[section_context_key(module, title)] = lines
    return result


def section_key(title: str) -> str:
    match = re.match(r"^([一二三四五六七八九十]+)、", title)
    if match:
        return match.group(1)
    return title


CIRCLED_NUMBERS = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
MODULE_TITLES = {"课文梳理", "知识梳理", "情境默写", "情景默写", "拓展阅读"}
MODULE_IMAGE_FILES = {
    "课文梳理": ("课文梳理.png", "wenyan_module_kewen.png"),
    "知识梳理": ("知识梳理.png", "wenyan_module_zhishi.png"),
    "情境默写": ("情景默写.png", "wenyan_module_moxie.png"),
    "情景默写": ("情景默写.png", "wenyan_module_moxie.png"),
    "拓展阅读": ("拓展阅读.png", "wenyan_module_tuozhan.png"),
}
ANIMATED_LINE_ROLES = {"answer", "analysis", "analysis_block", "segmentation_analysis", "yuedu_analysis"}


def canonical_module_title(title: str) -> str:
    return "情境默写" if title == "情景默写" else title


def split_sections_with_context(piece: Piece) -> list[tuple[str, str, list[str]]]:
    sections: list[tuple[str, str, list[str]]] = []
    current_module = ""
    for title, lines in split_sections(piece):
        if title in MODULE_TITLES:
            current_module = canonical_module_title(title)
        sections.append((current_module, title, lines))
    return sections


def section_context_key(module: str, title: str) -> str:
    return f"{module}\x1f{section_key(title)}"


def strip_inline_tags(text: str) -> str:
    return re.sub(r"<[^>]+>", "", html.unescape(text))


def clean_answer_piece(text: str) -> str:
    text = html.unescape(re.sub(r"<[^>]+>", "", text)).strip()
    text = re.sub(r"^[一二三四五六七八九十]+、\s*", "", text).strip()
    text = re.sub(r"^（?[一二三四五六七八九十]+）\s*", "", text).strip()
    text = re.sub(r"^\d+[\.、]\s*", "", text).strip()
    text = re.sub(r"^[%s]\s*" % re.escape(CIRCLED_NUMBERS), "", text).strip()
    text = text.strip("_＿ \t")
    return text.strip()


def split_answer_segment(segment: str) -> list[str]:
    parts = [clean_answer_piece(part) for part in re.split(r"[_＿]{2,}", segment)]
    return [part for part in parts if part]


def split_analysis(text: str) -> tuple[str, list[str]]:
    if "【解析】" not in text:
        return text, []
    before, after = text.split("【解析】", 1)
    analysis = f"【解析】{after.strip()}"
    return before.strip(), [analysis] if analysis.strip() != "【解析】" else []


def extract_ordered_answers(answer_lines: list[str] | None) -> list[str]:
    if not answer_lines:
        return []
    answers: list[str] = []
    circled_re = re.compile(f"[{re.escape(CIRCLED_NUMBERS)}]")

    for raw_line in answer_lines:
        line = raw_line.strip()
        if not line or line in MODULE_TITLES or line in VOLUMES:
            continue
        if re.fullmatch(r"[一二三四五六七八九十]+、", line):
            continue
        if line.startswith("【解析】"):
            continue
        line, _analyses = split_analysis(line)
        if not line:
            continue

        # Drop short labels such as "序：" or "正文：" before the real answer data.
        first_circled = circled_re.search(line)
        if "：" in line:
            label, rest = line.split("：", 1)
            if len(strip_inline_tags(label)) <= 8 and (first_circled or re.match(r"\s*\d+[\.、]", rest)):
                line = rest.strip()
                first_circled = circled_re.search(line)

        if first_circled:
            line = line[first_circled.start() :]
            matches = list(circled_re.finditer(line))
            for idx, match in enumerate(matches):
                start = match.end()
                end = matches[idx + 1].start() if idx + 1 < len(matches) else len(line)
                answers.extend(split_answer_segment(line[start:end]))
            continue

        answers.extend(split_answer_segment(line))
    return answers


def answer_text_to_values(text: str) -> list[str]:
    answer_part, _analyses = split_analysis(text)
    return extract_ordered_answers([answer_part])


def split_non_unique_answer_tail(answer_part: str) -> tuple[str, str | None]:
    if "答案不唯一" not in answer_part:
        return answer_part, None
    marker_match = re.search(r"[［\[]答案不唯一[］\]]", answer_part)
    marker_index = marker_match.start() if marker_match else len(answer_part)
    for index, ch in enumerate(answer_part[:marker_index]):
        if ch not in "（(":
            continue
        before = answer_part[:index]
        after = answer_part[index + 1 :]
        if re.search(r"[_＿]{2,}", before):
            tail = strip_inline_tags(re.sub(r"[_＿]{2,}", " ", after)).strip()
            tail = re.sub(r"\s+", " ", tail)
            return before.strip(), tail or None
    return answer_part, None


def append_answer_item_text(item: AnswerItem, text: str) -> None:
    answer_part, analyses = split_analysis(text.strip())
    optional_tail: str | None = None
    if answer_part:
        answer_part, optional_tail = split_non_unique_answer_tail(answer_part)
    if answer_part:
        item.answers.extend(answer_text_to_values(answer_part))
    if optional_tail:
        item.analyses.append(optional_tail)
    item.analyses.extend(analyses)


def looks_like_reference_translation_tail(line: str, current: AnswerItem | None) -> bool:
    if current is None or not current.answers:
        return False
    stripped = line.strip()
    if not stripped:
        return False
    plain = strip_inline_tags(stripped).strip()
    if not plain:
        return False
    if stripped.startswith("【解析】") or answer_item_match(stripped) or is_section_marker(stripped):
        return False
    if stripped in MODULE_TITLES or stripped in VOLUMES:
        return False
    if re.match(rf"^[{re.escape(CIRCLED_NUMBERS)}]", plain):
        return False
    if re.search(r"[_＿﹍]{2,}", stripped):
        return False
    if "参考译文" in plain or "译文" in plain[:12]:
        return True

    units = text_units(plain)
    if units < 30:
        return False
    prose_punctuation = len(re.findall(r"[。！？；：]", plain))
    answer_punctuation = sum(len(re.findall(r"[。！？；：]", answer)) for answer in current.answers)
    has_translation_cue = bool(re.search(r"[（(].{0,12}[他她它其我你]|说[：:]|曰[：:]|<span", stripped))
    return prose_punctuation >= 2 and (answer_punctuation >= 1 or has_translation_cue)


def is_section_marker(line: str) -> bool:
    return bool(re.fullmatch(r"[一二三四五六七八九十]+、", line.strip()))


def answer_item_match(line: str) -> re.Match[str] | None:
    # Supports both "1.xxx" and short labels such as "正文：1.xxx".
    return re.match(r"^(?:[^：:]{1,8}[：:])?\s*(\d+)[\.、]\s*(.*)$", line.strip())


def extract_answer_items(answer_lines: list[str] | None) -> dict[int, AnswerItem]:
    if not answer_lines:
        return {}
    items: dict[int, AnswerItem] = {}
    current: AnswerItem | None = None
    skipping_reference_tail = False
    for raw_line in answer_lines:
        line = raw_line.strip()
        if not line or line in MODULE_TITLES or line in VOLUMES or is_section_marker(line):
            skipping_reference_tail = False
            continue
        match = answer_item_match(line)
        if match:
            skipping_reference_tail = False
            current = AnswerItem(answers=[], analyses=[])
            items[int(match.group(1))] = current
            append_answer_item_text(current, match.group(2))
            continue
        if skipping_reference_tail:
            continue
        if looks_like_reference_translation_tail(line, current):
            skipping_reference_tail = True
            continue
        if current is not None:
            append_answer_item_text(current, line)
    return items


def extract_preamble_answers(answer_lines: list[str] | None) -> list[str]:
    if not answer_lines:
        return []
    preamble: list[str] = []
    for raw_line in answer_lines:
        line = raw_line.strip()
        if not line or line in MODULE_TITLES or line in VOLUMES or is_section_marker(line):
            continue
        if answer_item_match(line):
            break
        preamble.append(line)
    return extract_ordered_answers(preamble)


def extract_unitemized_analyses(answer_lines: list[str] | None) -> list[str]:
    if not answer_lines:
        return []
    analyses: list[str] = []
    for raw_line in answer_lines:
        line = raw_line.strip()
        if not line or line in MODULE_TITLES or line in VOLUMES or is_section_marker(line) or answer_item_match(line):
            continue
        _answer_part, line_analyses = split_analysis(line)
        analyses.extend(line_analyses)
    return analyses


def body_item_match(line: str) -> re.Match[str] | None:
    return re.match(r"^\s*(\d+)[\.、]\s*", strip_inline_tags(line).strip())


def split_body_numbered_groups(lines: list[str]) -> tuple[list[str], list[tuple[int, list[str]]]]:
    leading: list[str] = []
    groups: list[tuple[int, list[str]]] = []
    current_num: int | None = None
    current_lines: list[str] = []
    for line in lines:
        match = body_item_match(line)
        if match:
            if current_num is None:
                leading = current_lines
            else:
                groups.append((current_num, current_lines))
            current_num = int(match.group(1))
            current_lines = [line]
        else:
            current_lines.append(line)
    if current_num is None:
        leading = current_lines
    else:
        groups.append((current_num, current_lines))
    return leading, groups


LONG_ANSWER_UNDERLINE_CHARS = 62
LONG_ANSWER_LINE_TEXT_UNITS = 24

DOT_CHAR_SPACING = -864
DOT_BASELINE = -25000
DOT_SIZE_MULTIPLIER = 1.5
DOTTED_CHAR_SPACING = -1680
FIRST_LINE_INDENT_EMU = 24 * 12700 * 2
PARAGRAPH_LINE_SPACING_PCT = 118000
ANSWER_INLINE_VERTICAL_OFFSET_PT = 14
ANSWER_INLINE_VERTICAL_SHIFT_PT = 3.0
ANSWER_MULTILINE_VERTICAL_SHIFT_PT = 4.0
QUESTION_NUMBER_COLOR: str | None = None
INLINE_FILLED_ANSWER_ROLES = {"inline_answer_body"}


def is_long_answer_prompt(section_title: str, group_lines: list[str]) -> bool:
    start_index = 0
    for index, line in enumerate(group_lines):
        plain_line = strip_inline_tags(line).strip()
        if re.match(r"^\d+[\.．、]", plain_line) or line_has_fillable_blank(line):
            start_index = index
            break
    plain_group = "\n".join(strip_inline_tags(line) for line in group_lines[start_index:])
    triggers = [
        "翻译成现代汉语",
        "翻译",
        "简要说明",
        "简要分析",
        "论证方法",
        "请结合",
        "概括",
    ]
    return any(trigger in plain_group for trigger in triggers)


def should_embed_blank_answers(group_lines: list[str]) -> bool:
    plain_group = "\n".join(strip_inline_tags(line) for line in group_lines)
    return bool(re.search(r"正误关系|设错角度", plain_group))


def is_segmentation_prompt(group_lines: list[str]) -> bool:
    plain_group = "\n".join(strip_inline_tags(line) for line in group_lines)
    return "需要断句" in plain_group and "答案标号涂黑" in plain_group


def format_segmentation_text(line: str) -> str:
    plain = strip_inline_tags(line)
    plain = re.sub(r"\s+", "", plain)
    plain = re.sub(r"([A-H])", r" \1 ", plain)
    plain = re.sub(r"\s+", " ", plain).strip()
    return plain


def answer_letters(answer: str) -> str:
    return "".join(re.findall(r"[A-H]", strip_inline_tags(answer).upper()))


def segmentation_sentence_line(group_lines: list[str]) -> str:
    for line in group_lines[1:]:
        plain = strip_inline_tags(line).strip()
        if len(re.findall(r"[A-H]", plain)) >= 3:
            return format_segmentation_text(line)
    return format_segmentation_text(group_lines[-1] if group_lines else "")


def split_answer_for_fixed_lines(answer: str) -> list[str]:
    pieces: list[str] = []
    current = ""
    units = 0.0
    for ch in answer:
        weight = 1.0 if ord(ch) > 127 else 0.55
        if current and units + weight > LONG_ANSWER_LINE_TEXT_UNITS:
            pieces.append(current)
            current = ""
            units = 0.0
        current += ch
        units += weight
    if current:
        pieces.append(current)
    return pieces[:4]


def text_units(text: str) -> float:
    return sum(0.5 if ch == "\u00a0" else (1.0 if ord(ch) > 127 else 0.55) for ch in text)


def fixed_answer_line(answer: str = "") -> str:
    if not answer:
        return "_" * LONG_ANSWER_UNDERLINE_CHARS
    padded_answer = f"　{answer}　"
    padding = "_" * max(4, LONG_ANSWER_UNDERLINE_CHARS - int(math.ceil(text_units(padded_answer) * 2.0)))
    return answer_span(f"{padded_answer}{padding}", underline=True)


def fixed_long_answer_lines(answer: str) -> list[StyledLine]:
    pieces = split_answer_for_fixed_lines(answer)
    if not pieces:
        pieces = [""]
    line_count = max(2, len(pieces))
    lines = [
        StyledLine(fixed_answer_line(pieces[index] if index < len(pieces) else ""), "long_answer")
        for index in range(line_count)
    ]
    return lines


def answer_span(answer: str, *, underline: bool = False, pad_underline: bool = False) -> str:
    classes = "answer-fill uline" if underline else "answer-fill"
    text = html.escape(answer)
    if underline and pad_underline:
        text = f"　{text}　"
    return f'<span class="{classes}">{text}</span>'


def fill_parenthetical_blanks(line: str, answers: list[str], index: int) -> tuple[str, int, int]:
    filled = 0

    def repl(match: re.Match[str]) -> str:
        nonlocal index, filled
        if index >= len(answers):
            return match.group(0)
        answer = answers[index]
        index += 1
        filled += 1
        return f"{match.group(1)}{answer_span(answer)}{match.group(2)}"

    line = re.sub(r"([（(])\s*[_＿]{2,}\s*([）)])", repl, line)
    return line, index, filled


def fill_underlined_spans(line: str, answers: list[str], index: int) -> tuple[str, int, int]:
    filled = 0
    pattern = re.compile(r'<span\s+class="uline"([^>]*)>[\s_＿]+</span>')

    def repl(match: re.Match[str]) -> str:
        nonlocal index, filled
        if index >= len(answers):
            return match.group(0)
        answer = answers[index]
        index += 1
        filled += 1
        return answer_span(answer, underline=True, pad_underline=True)

    line = pattern.sub(repl, line)
    return line, index, filled


def looks_like_decorative_blank_line(line: str) -> bool:
    plain = strip_inline_tags(line).strip()
    if plain == "正____文":
        return True
    if re.match(r"^[_＿]{2,}（[^）]+）", plain):
        return True
    return False


def fill_plain_underline_blanks(line: str, answers: list[str], index: int) -> tuple[str, int, int]:
    if looks_like_decorative_blank_line(line):
        return line, index, 0
    filled = 0

    def repl(match: re.Match[str]) -> str:
        nonlocal index, filled
        if index >= len(answers):
            return match.group(0)
        answer = answers[index]
        index += 1
        filled += 1
        return answer_span(answer, underline=True, pad_underline=True)

    line = re.sub(r"[_＿]{3,}", repl, line)
    return line, index, filled


def line_has_fillable_blank(line: str) -> bool:
    if re.search(r"[（(]\s*[_＿]{2,}\s*[）)]", line):
        return True
    if re.search(r'<span\s+class="uline"[^>]*>[\s_＿]+</span>', line):
        return True
    if re.search(r"[_＿]{3,}", line) and not looks_like_decorative_blank_line(line):
        return True
    return False


def section_has_fillable_blanks(lines: list[str]) -> bool:
    return any(line_has_fillable_blank(line) for line in lines)


def fill_lines_with_answer_values(lines: list[str], answers: list[str]) -> tuple[list[str], int]:
    if not answers:
        return lines, 0
    index = 0
    filled_lines: list[str] = []
    for line in lines:
        line, index, filled = fill_parenthetical_blanks(line, answers, index)
        line, index, filled = fill_underlined_spans(line, answers, index)
        line, index, filled = fill_plain_underline_blanks(line, answers, index)
        filled_lines.append(line)
    return filled_lines, index


def strip_fillable_blanks_from_line(line: str) -> str:
    line = re.sub(r"[（(]\s*[_＿]{2,}\s*[）)]", "", line)
    line = re.sub(r'<span\s+class="uline"[^>]*>[\s_＿]+</span>', "", line)
    line = re.sub(r"[_＿]{3,}", "", line)
    return line.strip()


def is_piece_boundary_noise_line(line: str) -> bool:
    plain = strip_inline_tags(strip_fillable_blanks_from_line(line)).strip()
    compact = re.sub(r"[\s　]+", "", plain)
    if not compact:
        return False
    if compact in VOLUMES:
        return True
    for volume in VOLUMES:
        for prefix in (f"（{volume}）", f"({volume})"):
            if compact.startswith(prefix):
                tail = compact[len(prefix) :]
                return 0 < text_units(tail) <= 12 and not re.search(r"[，。！？；：、]", tail)
    return False


def remove_answer_blank_lines(lines: list[str]) -> list[str]:
    cleaned: list[str] = []
    for line in lines:
        if is_piece_boundary_noise_line(line):
            continue
        stripped = strip_fillable_blanks_from_line(line)
        if stripped:
            cleaned.append(stripped)
    return cleaned or lines


def embed_answers_in_lines(lines: list[str], answer_lines: list[str] | None) -> tuple[list[str], bool]:
    filled_lines, filled_count = fill_lines_with_answer_values(lines, extract_ordered_answers(answer_lines))
    return filled_lines, filled_count > 0


def estimate_line_count(line: str) -> int:
    line = re.sub(r"<img\b[^>]*>", "\u3000", line, flags=re.I)
    text = re.sub(r"<[^>]+>", "", line)
    text = html.unescape(text)
    units = 0
    for ch in text:
        units += 0.5 if ch == "\u00a0" else (1 if ord(ch) > 127 else 0.55)
    return max(1, math.ceil(units / 32))


OVERSIZED_BODY_FRAGMENT_UNITS = 280.0


def inline_run_to_html(run: InlineRun, text: str) -> str:
    if not text:
        return ""
    rendered = html.escape(text, quote=False)
    if run.superscript:
        rendered = f"<sup>{rendered}</sup>"
    classes: list[str] = []
    if run.answer_fill:
        classes.append("answer-fill")
    if run.underline == "wavy":
        classes.append("wave")
    elif run.underline:
        classes.append("uline")
    if run.underline_color == "00AEEF":
        classes.append("blue-underline")
    if run.highlight == "D9D9D9":
        classes.append("yuedu-bg-word")
    if run.emphasis_dot:
        classes.append("dot")
    if run.color == "00AEEF":
        classes.append("blue-text")
    if run.font == "宋体":
        classes.append("font-song")
    elif run.font in {"楷体", "KaiTi"}:
        classes.append("font-kaiti")
    elif run.font == "Times New Roman":
        classes.append("font-times")
    elif run.font == "黑体":
        classes.append("font-heiti")
    if run.bold:
        classes.append("bold-text")
    if classes:
        rendered = f'<span class="{" ".join(classes)}">{rendered}</span>'
    return rendered


def split_oversized_styled_line(line: StyledLine, max_units: float = OVERSIZED_BODY_FRAGMENT_UNITS) -> list[StyledLine]:
    if line.role not in {"body", "reading", "extended_reading", "deep_material_reading"}:
        return [line]
    if estimate_line_count(line.text) <= 10:
        return [line]

    fragments: list[str] = []
    current_parts: list[str] = []
    current_units = 0.0

    for run in parse_inline_runs(line.text):
        if run.tif_name:
            token = f'<img class="fbd-tif-char" src="{html.escape(run.tif_name, quote=True)}" alt="{html.escape(run.tif_name, quote=True)}" data-fbd-tif="{html.escape(run.tif_name, quote=True)}" />'
            weight = 1.0
            if current_parts and current_units + weight > max_units:
                fragments.append("".join(current_parts).strip())
                current_parts = []
                current_units = 0.0
            current_parts.append(token)
            current_units += weight
            continue
        if run.answer_fill:
            token = inline_run_to_html(run, run.text)
            weight = text_units(run.text)
            if current_parts and current_units + weight > max_units:
                fragments.append("".join(current_parts).strip())
                current_parts = []
                current_units = 0.0
            current_parts.append(token)
            current_units += weight
            continue
        for ch in run.text:
            weight = 1.0 if ord(ch) > 127 else 0.55
            if current_parts and current_units + weight > max_units:
                fragments.append("".join(current_parts).strip())
                current_parts = []
                current_units = 0.0
            current_parts.append(inline_run_to_html(run, ch))
            current_units += weight

    if current_parts:
        fragments.append("".join(current_parts).strip())
    clean_fragments = [fragment for fragment in fragments if fragment]
    if len(clean_fragments) <= 1:
        return [line]
    split_lines: list[StyledLine] = []
    force_first_new_chunk = bool(re.match(r"^\s*\d+[\.．、]", strip_inline_tags(line.text)))
    for index, fragment in enumerate(clean_fragments):
        meta = dict(line.meta or {})
        if force_first_new_chunk and index == 0:
            meta["force_new_chunk"] = True
        split_lines.append(StyledLine(fragment, line.role, meta))
    return split_lines


def is_choice_option_line(line: StyledLine) -> bool:
    plain = strip_inline_tags(line.text).strip()
    return bool(re.match(r"^[A-D][\.．、]", plain))


def is_numbered_question_line(line: StyledLine) -> bool:
    if line.role not in {"body", "reading", "extended_reading"}:
        return False
    return bool(re.match(r"^\s*\d+[\.．、]\s*", strip_inline_tags(line.text).strip()))


def line_page_cost(line: StyledLine) -> int:
    cost = estimate_line_count(line.text)
    if line.role in {
        "heading",
        "module_header",
        "module",
        "answer_heading",
        "subheading",
        "center_heading",
        "segmentation_question",
    }:
        cost += 1
    return cost


def block_page_cost(block: LineBlock) -> int:
    return sum(line_page_cost(line) for line in block.lines)


def is_choice_body_block(lines: list[StyledLine]) -> bool:
    if not lines or not is_numbered_question_line(lines[0]):
        return False
    letters = {
        strip_inline_tags(line.text).strip()[0]
        for line in lines[1:]
        if is_choice_option_line(line)
    }
    return {"A", "B", "C", "D"}.issubset(letters)


def logical_line_blocks(lines: list[StyledLine]) -> list[LineBlock]:
    blocks: list[LineBlock] = []
    index = 0
    while index < len(lines):
        line = lines[index]
        if is_numbered_question_line(line):
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
                if is_numbered_question_line(next_line):
                    break
                end += 1
            body_lines = lines[index:end]
            if is_choice_body_block(body_lines):
                blocks.append(LineBlock(body_lines, "choice_body"))
                analysis_end = end
                while analysis_end < len(lines) and lines[analysis_end].role in {"analysis", "analysis_block"}:
                    analysis_end += 1
                if analysis_end > end:
                    blocks.append(LineBlock(lines[end:analysis_end], "choice_analysis"))
                index = analysis_end
                continue
        blocks.append(LineBlock([line], "single"))
        index += 1
    return blocks


def append_block_to_current(current: list[StyledLine], block: LineBlock) -> None:
    current.extend(block.lines)


def chunk_styled_lines(lines: list[StyledLine], max_units: int = 12) -> list[list[StyledLine]]:
    expanded_lines: list[StyledLine] = []
    for line in lines:
        expanded_lines.extend(split_oversized_styled_line(line))
    blocks = logical_line_blocks(expanded_lines)

    chunks: list[list[StyledLine]] = []
    current: list[StyledLine] = []
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
        cost = block_page_cost(block)
        if block.kind == "choice_body":
            next_block = blocks[index + 1] if index + 1 < len(blocks) and blocks[index + 1].kind == "choice_analysis" else None
            next_cost = block_page_cost(next_block) if next_block else 0
            preferred_cost = cost + next_cost if next_block and cost + next_cost <= max_units else cost
            if current and (used + preferred_cost > max_units or used + cost > max_units):
                chunks.append(current)
                current = []
                used = 0
            append_block_to_current(current, block)
            used += cost
            if next_block:
                if used + next_cost > max_units and current:
                    chunks.append(current)
                    current = []
                    used = 0
                append_block_to_current(current, next_block)
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
                used = line_page_cost(carry)
            else:
                chunks.append(current)
                current = []
                used = 0
        append_block_to_current(current, block)
        used += cost
        index += 1
    if current:
        chunks.append(current)
    return chunks


def canonical_center_heading(line: str) -> str | None:
    plain = strip_inline_tags(line).strip()
    compact = re.sub(r"[_＿\s]+", "", plain)
    if compact == "序":
        return "序"
    if compact == "正文":
        return "正　文"
    return None


def is_reading_section(section_title: str) -> bool:
    plain = strip_inline_tags(section_title)
    return "请解释加点字词的意思" in plain


def is_subheading_line(lines: list[str], index: int) -> bool:
    line = strip_inline_tags(lines[index]).strip()
    match = re.match(r"^\d+[\.、]\s*(.+)$", line)
    if not match:
        return False
    label = match.group(1).strip()
    if text_units(label) > 8 or re.search(r"[，。；：？！、“”（）《》]", label):
        return False
    for next_line in lines[index + 1 :]:
        plain_next = strip_inline_tags(next_line).strip()
        if not plain_next:
            continue
        return bool(re.match(rf"^[{re.escape(CIRCLED_NUMBERS)}]", plain_next))
    return False


def is_source_note_plain(plain: str) -> bool:
    return bool(re.match(r"^[\uFF08(].*(\u8282\u9009\u81ea|\u6458\u7f16\u81ea|\u9009\u81ea)", plain))


def styled_body_lines(lines: list[str], section_title: str = "") -> list[StyledLine]:
    styled: list[StyledLine] = []
    reading = is_reading_section(section_title)
    plain_lines = [strip_inline_tags(line).strip() for line in lines]
    extended_reading_start = next((i for i, plain in enumerate(plain_lines) if plain.startswith("阅读下面")), None)
    for index, line in enumerate(lines):
        plain = plain_lines[index]
        if line in MODULE_TITLES or line in VOLUMES or is_piece_boundary_noise_line(line):
            continue
        center_text = canonical_center_heading(line)
        if center_text:
            styled.append(StyledLine(center_text, "center_heading"))
            continue
        if re.match(r"^[一二三四五六七八九十]+、", line):
            styled.append(StyledLine(line, "heading"))
        elif is_source_note_plain(plain):
            styled.append(StyledLine(line, "source_note"))
        elif plain.startswith("【注】"):
            styled.append(StyledLine(line, "annotation"))
        elif is_subheading_line(lines, index):
            styled.append(StyledLine(line, "subheading"))
        elif reading:
            styled.append(StyledLine(line, "reading"))
        elif (
            extended_reading_start is not None
            and index > extended_reading_start
            and plain
            and not re.match(r"^\d+[\.．、]", plain)
            and not re.match(r"^[A-D][\.．、]", plain)
        ):
            styled.append(StyledLine(line, "extended_reading"))
        else:
            styled.append(StyledLine(line, "body"))
    return styled


def build_styled_section(lines: list[str], answer_lines: list[str] | None) -> tuple[list[StyledLine], bool]:
    section_title = lines[0] if lines else ""
    answer_items = extract_answer_items(answer_lines)
    ordered_answers = extract_ordered_answers(answer_lines)
    ordered_answer_index = 0
    leading, groups = split_body_numbered_groups(lines)
    if answer_items and groups:
        styled: list[StyledLine] = []
        preamble_lines, preamble_filled = fill_lines_with_answer_values(leading, extract_preamble_answers(answer_lines))
        preamble_styled = styled_body_lines(preamble_lines, section_title)
        styled.extend(preamble_styled)
        matched_any = preamble_filled

        for num, group_lines in groups:
            item = answer_items.get(num)
            if item is None:
                if section_has_fillable_blanks(group_lines) and ordered_answer_index < len(ordered_answers):
                    filled_group, group_filled = fill_lines_with_answer_values(
                        group_lines,
                        ordered_answers[ordered_answer_index:],
                    )
                    ordered_answer_index += group_filled
                    styled.extend(styled_body_lines(filled_group, section_title))
                    matched_any = matched_any or bool(group_filled)
                    continue
                styled.extend(styled_body_lines(group_lines, section_title))
                continue

            if is_segmentation_prompt(group_lines):
                question_line = group_lines[0] if group_lines else ""
                answer = answer_letters("".join(item.answers))
                styled.append(StyledLine(strip_fillable_blanks_from_line(question_line), "segmentation_question"))
                styled.append(
                    StyledLine(
                        segmentation_sentence_line(group_lines),
                        "segmentation_text",
                        {"answer": answer},
                    )
                )
                for analysis in item.analyses:
                    styled.append(StyledLine(analysis, "segmentation_analysis"))
                styled.append(StyledLine("", "page_break"))
                matched_any = True
                continue
            elif item.answers and is_long_answer_prompt(section_title, group_lines) and not should_embed_blank_answers(group_lines):
                question_lines = remove_answer_blank_lines(group_lines)
                styled.extend(styled_body_lines(question_lines, section_title))
                styled.extend(fixed_long_answer_lines("；".join(item.answers)))
                if item.analyses:
                    styled.append(StyledLine("", "answer_spacer"))
                matched_any = True
            elif section_has_fillable_blanks(group_lines):
                fill_source = item.answers
                if not fill_source and ordered_answer_index < len(ordered_answers):
                    fill_source = ordered_answers[ordered_answer_index:]
                filled_group, group_filled = fill_lines_with_answer_values(group_lines, fill_source)
                if not item.answers and group_filled:
                    ordered_answer_index += group_filled
                styled.extend(styled_body_lines(filled_group, section_title))
                matched_any = matched_any or bool(group_filled)
            else:
                styled.extend(styled_body_lines(group_lines, section_title))
                if item.answers and is_long_answer_prompt(section_title, group_lines):
                    styled.extend(fixed_long_answer_lines("；".join(item.answers)))
                    if item.analyses:
                        styled.append(StyledLine("", "answer_spacer"))
                else:
                    for answer in item.answers:
                        styled.append(StyledLine(answer, "answer"))
                matched_any = matched_any or bool(item.answers)

            for analysis in item.analyses:
                styled.append(StyledLine(analysis, "analysis"))
                matched_any = True
        return styled, matched_any

    if section_has_fillable_blanks(lines):
        filled_lines, embedded_answers = embed_answers_in_lines(lines, answer_lines)
        styled = styled_body_lines(filled_lines, section_title)
        if embedded_answers:
            for analysis in extract_unitemized_analyses(answer_lines):
                styled.append(StyledLine(analysis, "analysis"))
        return styled, embedded_answers
    return styled_body_lines(lines, section_title), False


def make_piece_slide_chunks(body_piece: Piece, answer_piece: Piece | None) -> list[list[StyledLine]]:
    answer_map = answer_section_map(answer_piece)
    chunks: list[list[StyledLine]] = []
    pending_module_header: str | None = None
    for module, title, lines in split_sections_with_context(body_piece):
        if title in MODULE_TITLES:
            pending_module_header = canonical_module_title(title)
            lines = [line for line in lines[1:] if line.strip()]
            if not lines:
                continue
        answer_lines = answer_map.get(section_context_key(module, title))
        if not answer_lines and module:
            module_answer_lines = answer_map.get(section_context_key(module, module))
            if module_answer_lines and any("【解析】" in line for line in module_answer_lines):
                answer_lines = module_answer_lines
        styled, _matched_answers = build_styled_section(lines, answer_lines)
        if pending_module_header:
            styled = [StyledLine(pending_module_header, "module_header")] + styled
            pending_module_header = None
        chunks.extend(chunk_styled_lines(styled))
    return chunks


def parse_inline_runs(text: str) -> list[InlineRun]:
    parser = InlineMarkdownParser()
    try:
        parser.feed(text)
        parser.close()
    except Exception:
        return [InlineRun(re.sub(r"<[^>]+>", "", text))]
    return parser.runs or [InlineRun(re.sub(r"<[^>]+>", "", text))]


def clear_text_shapes(root: ET.Element, keep_substrings: Iterable[str] = ()) -> None:
    keep = tuple(keep_substrings)
    for sp in root.findall(f".//{{{P_NS}}}sp"):
        texts = sp.findall(f".//{{{A_NS}}}t")
        joined = "".join(t.text or "" for t in texts)
        if keep and any(part in joined for part in keep):
            continue
        for t in texts:
            t.text = ""


def get_shape_text(sp: ET.Element) -> str:
    return "".join(t.text or "" for t in sp.findall(f".//{{{A_NS}}}t"))


def set_shape_text(sp: ET.Element, text: str) -> bool:
    texts = sp.findall(f".//{{{A_NS}}}t")
    if not texts:
        return False
    texts[0].text = text
    for t in texts[1:]:
        t.text = ""
    return True


def style_existing_text_shape(sp: ET.Element, *, font: str, color: str, bold: bool | None = None) -> None:
    for r_pr in sp.findall(f".//{{{A_NS}}}rPr"):
        if bold is not None:
            if bold:
                r_pr.attrib["b"] = "1"
            else:
                r_pr.attrib.pop("b", None)
        for child in list(r_pr):
            if child.tag in {
                qn(A_NS, "solidFill"),
                qn(A_NS, "latin"),
                qn(A_NS, "ea"),
                qn(A_NS, "cs"),
            }:
                r_pr.remove(child)
        solid = ET.Element(qn(A_NS, "solidFill"))
        ET.SubElement(solid, qn(A_NS, "srgbClr"), {"val": color})
        r_pr.insert(0, solid)
        ET.SubElement(r_pr, qn(A_NS, "latin"), {"typeface": font})
        ET.SubElement(r_pr, qn(A_NS, "ea"), {"typeface": font})
        ET.SubElement(r_pr, qn(A_NS, "cs"), {"typeface": font})


def text_shapes(root: ET.Element) -> list[ET.Element]:
    return [sp for sp in root.findall(f".//{{{P_NS}}}sp") if sp.findall(f".//{{{A_NS}}}t")]


def replace_exact_text(root: ET.Element, old: str, new: str) -> bool:
    replaced = False
    for sp in text_shapes(root):
        if get_shape_text(sp) == old:
            set_shape_text(sp, new)
            replaced = True
    return replaced


def remove_descendant(root: ET.Element, target: ET.Element) -> bool:
    for child in list(root):
        if child is target:
            root.remove(child)
            return True
        if remove_descendant(child, target):
            return True
    return False


def strip_copied_custom_data(root: ET.Element) -> None:
    """Remove duplicated PowerPoint slide tags copied from template slides.

    Template slides contain p:custDataLst/p:tags records. A single copied slide
    is fine, but many generated slides pointing to the same tag part can make
    PowerPoint refuse to open the deck. These records are authoring metadata,
    not visible content.
    """

    for elem in list(root.iter()):
        if elem.tag == qn(P_NS, "custDataLst"):
            remove_descendant(root, elem)


def next_shape_id(root: ET.Element) -> int:
    ids = []
    for c_nv_pr in root.findall(f".//{{{P_NS}}}cNvPr"):
        value = c_nv_pr.attrib.get("id")
        if value and value.isdigit():
            ids.append(int(value))
    return (max(ids) + 1) if ids else 1000


def run_properties(
    font: str,
    size_pt: float,
    color: str,
    bold: bool,
    underline: str | None,
    superscript: bool,
    italic: bool = False,
    underline_color: str | None = None,
    highlight: str | None = None,
    *,
    char_spacing: int | None = None,
    baseline: int | None = None,
    alpha: int | None = None,
) -> ET.Element:
    attrs = {"lang": "zh-CN", "sz": str(int(size_pt * 100))}
    if bold:
        attrs["b"] = "1"
    if italic:
        attrs["i"] = "1"
    if underline:
        attrs["u"] = underline
    if char_spacing is not None:
        attrs["spc"] = str(char_spacing)
    if superscript:
        attrs["baseline"] = "30000"
        attrs["sz"] = str(int(size_pt * 72))
    if baseline is not None:
        attrs["baseline"] = str(baseline)
    rpr = ET.Element(qn(A_NS, "rPr"), attrs)
    solid = ET.SubElement(rpr, qn(A_NS, "solidFill"))
    srgb = ET.SubElement(solid, qn(A_NS, "srgbClr"), {"val": color})
    if alpha is not None:
        ET.SubElement(srgb, qn(A_NS, "alpha"), {"val": str(alpha)})
    if highlight:
        highlight_el = ET.SubElement(rpr, qn(A_NS, "highlight"))
        ET.SubElement(highlight_el, qn(A_NS, "srgbClr"), {"val": highlight})
    if underline and underline_color:
        u_fill = ET.SubElement(rpr, qn(A_NS, "uFill"))
        u_solid = ET.SubElement(u_fill, qn(A_NS, "solidFill"))
        ET.SubElement(u_solid, qn(A_NS, "srgbClr"), {"val": underline_color})
    ET.SubElement(rpr, qn(A_NS, "latin"), {"typeface": font})
    ET.SubElement(rpr, qn(A_NS, "ea"), {"typeface": font})
    ET.SubElement(rpr, qn(A_NS, "cs"), {"typeface": font})
    return rpr


def is_emphasis_dot_char(ch: str) -> bool:
    return bool(re.match(r"[\u4e00-\u9fffA-Za-z0-9]", ch))


def append_run(
    paragraph: ET.Element,
    text: str,
    *,
    font: str,
    size: float,
    color: str,
    bold: bool,
    underline: str | None = None,
    superscript: bool = False,
    italic: bool = False,
    underline_color: str | None = None,
    highlight: str | None = None,
    char_spacing: int | None = None,
    baseline: int | None = None,
    alpha: int | None = None,
) -> None:
    if not text:
        return
    r = ET.SubElement(paragraph, qn(A_NS, "r"))
    r.append(
        run_properties(
            font,
            size,
            color,
            bold,
            underline,
            superscript,
            italic,
            underline_color,
            highlight,
            char_spacing=char_spacing,
            baseline=baseline,
            alpha=alpha,
        )
    )
    t = ET.SubElement(r, qn(A_NS, "t"))
    if text[:1].isspace() or text[-1:].isspace() or "  " in text or "　" in text:
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text


def append_emphasis_dot_runs(
    paragraph: ET.Element,
    text: str,
    *,
    font: str,
    size: float,
    color: str,
    bold: bool,
    underline: str | None,
    superscript: bool,
    italic: bool = False,
    underline_color: str | None = None,
    highlight: str | None = None,
    alpha: int | None = None,
) -> None:
    normal_buf: list[str] = []

    def flush_normal() -> None:
        if normal_buf:
            append_run(
                paragraph,
                "".join(normal_buf),
                font=font,
                size=size,
                color=color,
                bold=bold,
                underline=underline,
                superscript=superscript,
                italic=italic,
                underline_color=underline_color,
                highlight=highlight,
                alpha=alpha,
            )
            normal_buf.clear()

    for ch in text:
        if not is_emphasis_dot_char(ch):
            normal_buf.append(ch)
            continue
        flush_normal()
        append_run(
            paragraph,
            ch,
            font=font,
            size=size,
            color=color,
            bold=bold,
            underline=underline,
            superscript=superscript,
            italic=italic,
            underline_color=underline_color,
            highlight=highlight,
            char_spacing=DOTTED_CHAR_SPACING,
            alpha=alpha,
        )
        append_run(
            paragraph,
            "．",
            font=font,
            size=size * DOT_SIZE_MULTIPLIER,
            color=color,
            bold=False,
            underline=None,
            superscript=False,
            italic=italic,
            char_spacing=DOT_CHAR_SPACING,
            baseline=DOT_BASELINE,
            alpha=alpha,
        )
    flush_normal()


def paragraph_properties(role: str) -> dict[str, str]:
    attrs = {
        "marL": "0",
        "indent": "0",
        "algn": "l",
        "eaLnBrk": "1",
        "latinLnBrk": "0",
        "hangingPunct": "0",
    }
    if role in {"body", "reading", "long_answer", "extended_reading", "inline_answer_body"}:
        attrs["indent"] = str(FIRST_LINE_INDENT_EMU)
    if role == "analysis":
        attrs["marL"] = str(FIRST_LINE_INDENT_EMU)
    if role == "reading":
        attrs["algn"] = "just"
    if role == "extended_reading":
        attrs["algn"] = "just"
    if role == "source_note":
        attrs["algn"] = "r"
    if role == "center_heading":
        attrs["algn"] = "ctr"
    return attrs


def add_paragraph_line_spacing(p_pr: ET.Element) -> None:
    ln_spc = ET.SubElement(p_pr, qn(A_NS, "lnSpc"))
    ET.SubElement(ln_spc, qn(A_NS, "spcPct"), {"val": str(PARAGRAPH_LINE_SPACING_PCT)})


def append_to_sp_tree(root: ET.Element, sp: ET.Element) -> None:
    sp_tree = root.find(f".//{{{P_NS}}}spTree")
    if sp_tree is None:
        raise ValueError("slide spTree not found")
    ext_lst = sp_tree.find(qn(P_NS, "extLst"))
    if ext_lst is not None:
        index = list(sp_tree).index(ext_lst)
        sp_tree.insert(index, sp)
    else:
        sp_tree.append(sp)


def next_rel_id(rels_root: ET.Element) -> str:
    max_id = 0
    for rel in rels_root:
        rel_id = rel.attrib.get("Id", "")
        match = re.match(r"rId(\d+)$", rel_id)
        if match:
            max_id = max(max_id, int(match.group(1)))
    return f"rId{max_id + 1}"


def add_image_relationship(rels_root: ET.Element, target: str) -> str:
    for rel in rels_root:
        if rel.attrib.get("Type") == REL_IMAGE_TYPE and rel.attrib.get("Target") == target:
            return rel.attrib["Id"]
    rel_id = next_rel_id(rels_root)
    ET.SubElement(
        rels_root,
        qn(REL_NS, "Relationship"),
        {"Id": rel_id, "Type": REL_IMAGE_TYPE, "Target": target},
    )
    return rel_id


def add_slide_jump_relationship(rels_root: ET.Element, slide_number: int) -> str:
    target = f"slide{slide_number}.xml"
    for rel in rels_root:
        if rel.attrib.get("Type") == REL_SLIDE_TYPE and rel.attrib.get("Target") == target:
            return rel.attrib["Id"]
    rel_id = next_rel_id(rels_root)
    ET.SubElement(
        rels_root,
        qn(REL_NS, "Relationship"),
        {"Id": rel_id, "Type": REL_SLIDE_TYPE, "Target": target},
    )
    return rel_id


def append_picture(
    root: ET.Element,
    *,
    rel_id: str,
    x: float,
    y: float,
    w: float,
    h: float,
    name: str,
) -> int:
    shape_id = next_shape_id(root)
    pic = ET.Element(qn(P_NS, "pic"))
    nv_pic_pr = ET.SubElement(pic, qn(P_NS, "nvPicPr"))
    ET.SubElement(nv_pic_pr, qn(P_NS, "cNvPr"), {"id": str(shape_id), "name": f"{name} {shape_id}"})
    c_nv_pic_pr = ET.SubElement(nv_pic_pr, qn(P_NS, "cNvPicPr"))
    ET.SubElement(c_nv_pic_pr, qn(A_NS, "picLocks"), {"noChangeAspect": "1"})
    ET.SubElement(nv_pic_pr, qn(P_NS, "nvPr"))

    blip_fill = ET.SubElement(pic, qn(P_NS, "blipFill"))
    ET.SubElement(blip_fill, qn(A_NS, "blip"), {qn(R_NS, "embed"): rel_id})
    stretch = ET.SubElement(blip_fill, qn(A_NS, "stretch"))
    ET.SubElement(stretch, qn(A_NS, "fillRect"))

    sp_pr = ET.SubElement(pic, qn(P_NS, "spPr"))
    xfrm = ET.SubElement(sp_pr, qn(A_NS, "xfrm"))
    ET.SubElement(xfrm, qn(A_NS, "off"), {"x": str(inch(x)), "y": str(inch(y))})
    ET.SubElement(xfrm, qn(A_NS, "ext"), {"cx": str(inch(w)), "cy": str(inch(h))})
    geom = ET.SubElement(sp_pr, qn(A_NS, "prstGeom"), {"prst": "rect"})
    ET.SubElement(geom, qn(A_NS, "avLst"))
    append_to_sp_tree(root, pic)
    return shape_id


def append_dashed_line(
    root: ET.Element,
    *,
    x: float,
    y: float,
    w: float,
    name: str = "generated dashed line",
) -> int:
    shape_id = next_shape_id(root)
    cxn = ET.Element(qn(P_NS, "cxnSp"))
    nv_pr = ET.SubElement(cxn, qn(P_NS, "nvCxnSpPr"))
    ET.SubElement(nv_pr, qn(P_NS, "cNvPr"), {"id": str(shape_id), "name": f"{name} {shape_id}"})
    ET.SubElement(nv_pr, qn(P_NS, "cNvCxnSpPr"))
    ET.SubElement(nv_pr, qn(P_NS, "nvPr"))

    sp_pr = ET.SubElement(cxn, qn(P_NS, "spPr"))
    xfrm = ET.SubElement(sp_pr, qn(A_NS, "xfrm"))
    ET.SubElement(xfrm, qn(A_NS, "off"), {"x": str(inch(x)), "y": str(inch(y))})
    ET.SubElement(xfrm, qn(A_NS, "ext"), {"cx": str(inch(w)), "cy": "0"})
    geom = ET.SubElement(sp_pr, qn(A_NS, "prstGeom"), {"prst": "line"})
    ET.SubElement(geom, qn(A_NS, "avLst"))
    ln = ET.SubElement(sp_pr, qn(A_NS, "ln"), {"w": "9525"})
    solid = ET.SubElement(ln, qn(A_NS, "solidFill"))
    ET.SubElement(solid, qn(A_NS, "srgbClr"), {"val": "000000"})
    ET.SubElement(ln, qn(A_NS, "prstDash"), {"val": "dash"})
    append_to_sp_tree(root, cxn)
    return shape_id


def module_image_package(module_title: str) -> tuple[Path, str] | None:
    info = MODULE_IMAGE_FILES.get(module_title)
    if not info:
        return None
    file_name, package_name = info
    path = MODULE_IMAGE_DIR / file_name
    if not path.exists():
        return None
    return path, f"ppt/media/{package_name}"


def add_module_header_image(root: ET.Element, rels_root: ET.Element, module_title: str) -> None:
    image_info = module_image_package(module_title)
    if image_info is None:
        return
    _image_path, package_path = image_info
    rel_id = add_image_relationship(rels_root, posixpath.relpath(package_path, "ppt/slides"))
    image_w = 2.45
    image_h = image_w * 69 / 270
    append_picture(
        root,
        rel_id=rel_id,
        x=(SLIDE_W_IN - image_w) / 2,
        y=0.18,
        w=image_w,
        h=image_h,
        name=f"{module_title} header",
    )


def image_to_png_bytes(path: Path) -> bytes:
    from PIL import Image

    with Image.open(path) as image:
        if image.mode not in {"RGB", "RGBA"}:
            image = image.convert("RGBA")
        buffer = io.BytesIO()
        image.save(buffer, format="PNG")
        return buffer.getvalue()


def tif_package_path(path: Path) -> str:
    digest = hashlib.sha1(str(path).encode("utf-8")).hexdigest()[:10]
    safe_stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", path.stem).strip("_") or "image"
    return f"ppt/media/fbd_tif_{digest}_{safe_stem}.png"


def resolve_tif_path(tif_name: str, search_dirs: list[Path]) -> Path | None:
    wanted = tif_name.lower()
    for directory in search_dirs:
        direct = directory / tif_name
        if direct.exists():
            return direct
        if not directory.exists():
            continue
        for candidate in directory.rglob("*.tif"):
            if candidate.name.lower() == wanted:
                return candidate
        for candidate in directory.rglob("*.tiff"):
            if candidate.name.lower() == wanted:
                return candidate
    return None


def append_tif_picture(
    root: ET.Element,
    rels_root: ET.Element,
    media_assets: dict[str, bytes],
    *,
    tif_path: Path,
    x: float,
    y: float,
    h: float,
    name: str,
) -> int:
    package_path = tif_package_path(tif_path)
    if package_path not in media_assets:
        media_assets[package_path] = image_to_png_bytes(tif_path)
    rel_id = add_image_relationship(rels_root, posixpath.relpath(package_path, "ppt/slides"))

    try:
        from PIL import Image

        with Image.open(tif_path) as image:
            width, height = image.size
        aspect = width / height if height else 1.0
    except Exception:
        aspect = 1.0
    return append_picture(
        root,
        rel_id=rel_id,
        x=x,
        y=y,
        w=max(0.08, h * aspect),
        h=h,
        name=name,
    )


def append_rect(
    root: ET.Element,
    *,
    x: float,
    y: float,
    w: float,
    h: float,
    fill_color: str | None = None,
    line_color: str | None = "000000",
    name: str = "generated rect",
) -> int:
    shape_id = next_shape_id(root)
    sp = ET.Element(qn(P_NS, "sp"))
    nv_sp_pr = ET.SubElement(sp, qn(P_NS, "nvSpPr"))
    ET.SubElement(nv_sp_pr, qn(P_NS, "cNvPr"), {"id": str(shape_id), "name": f"{name} {shape_id}"})
    ET.SubElement(nv_sp_pr, qn(P_NS, "cNvSpPr"))
    ET.SubElement(nv_sp_pr, qn(P_NS, "nvPr"))

    sp_pr = ET.SubElement(sp, qn(P_NS, "spPr"))
    xfrm = ET.SubElement(sp_pr, qn(A_NS, "xfrm"))
    ET.SubElement(xfrm, qn(A_NS, "off"), {"x": str(inch(x)), "y": str(inch(y))})
    ET.SubElement(xfrm, qn(A_NS, "ext"), {"cx": str(inch(w)), "cy": str(inch(h))})
    geom = ET.SubElement(sp_pr, qn(A_NS, "prstGeom"), {"prst": "rect"})
    ET.SubElement(geom, qn(A_NS, "avLst"))
    if fill_color:
        fill = ET.SubElement(sp_pr, qn(A_NS, "solidFill"))
        ET.SubElement(fill, qn(A_NS, "srgbClr"), {"val": fill_color})
    else:
        ET.SubElement(sp_pr, qn(A_NS, "noFill"))
    ln = ET.SubElement(sp_pr, qn(A_NS, "ln"))
    if line_color:
        solid = ET.SubElement(ln, qn(A_NS, "solidFill"))
        ET.SubElement(solid, qn(A_NS, "srgbClr"), {"val": line_color})
    else:
        ET.SubElement(ln, qn(A_NS, "noFill"))

    tx_body = ET.SubElement(sp, qn(P_NS, "txBody"))
    ET.SubElement(tx_body, qn(A_NS, "bodyPr"), {"rtlCol": "0"})
    ET.SubElement(tx_body, qn(A_NS, "lstStyle"))
    ET.SubElement(tx_body, qn(A_NS, "p"))
    append_to_sp_tree(root, sp)
    return shape_id


def segmentation_char_width(ch: str) -> float:
    if ch == " ":
        return 0.15
    if re.match(r"[A-H]", ch):
        return 0.20
    if ord(ch) > 127:
        return 0.33
    return 0.16


def segmentation_letter_boxes(sentence: str, x: float) -> list[tuple[str, float]]:
    boxes: list[tuple[str, float]] = []
    cursor = 0.0
    for ch in sentence:
        if re.match(r"[A-H]", ch):
            boxes.append((ch, x + cursor - 0.04))
        cursor += segmentation_char_width(ch)
    return boxes


def segmentation_text_width(sentence: str) -> float:
    return sum(segmentation_char_width(ch) for ch in sentence)


def append_textbox(
    root: ET.Element,
    *,
    x: float,
    y: float,
    w: float,
    h: float,
    lines: list[StyledLine],
    default_role: str = "body",
    name: str = "generated text",
    answer_mode: str = "normal",
    target_event_index: int | None = None,
) -> int:
    if not lines:
        lines = [StyledLine("", default_role)]
    sp_tree = root.find(f".//{{{P_NS}}}spTree")
    if sp_tree is None:
        raise ValueError("slide spTree not found")

    shape_id = next_shape_id(root)
    sp = ET.Element(qn(P_NS, "sp"))

    nv_sp_pr = ET.SubElement(sp, qn(P_NS, "nvSpPr"))
    ET.SubElement(nv_sp_pr, qn(P_NS, "cNvPr"), {"id": str(shape_id), "name": f"{name} {shape_id}"})
    ET.SubElement(nv_sp_pr, qn(P_NS, "cNvSpPr"), {"txBox": "1"})
    ET.SubElement(nv_sp_pr, qn(P_NS, "nvPr"))

    sp_pr = ET.SubElement(sp, qn(P_NS, "spPr"))
    xfrm = ET.SubElement(sp_pr, qn(A_NS, "xfrm"))
    ET.SubElement(xfrm, qn(A_NS, "off"), {"x": str(inch(x)), "y": str(inch(y))})
    ET.SubElement(xfrm, qn(A_NS, "ext"), {"cx": str(inch(w)), "cy": str(inch(h))})
    geom = ET.SubElement(sp_pr, qn(A_NS, "prstGeom"), {"prst": "rect"})
    ET.SubElement(geom, qn(A_NS, "avLst"))
    ET.SubElement(sp_pr, qn(A_NS, "noFill"))
    ln = ET.SubElement(sp_pr, qn(A_NS, "ln"))
    ET.SubElement(ln, qn(A_NS, "noFill"))

    tx_body = ET.SubElement(sp, qn(P_NS, "txBody"))
    ET.SubElement(
        tx_body,
        qn(A_NS, "bodyPr"),
        {
            "wrap": "square",
            "rtlCol": "0",
            "lIns": "0",
            "tIns": "0",
            "rIns": "0",
            "bIns": "0",
        },
    )
    ET.SubElement(tx_body, qn(A_NS, "lstStyle"))

    event_index = 0
    active_long_answer_event: int | None = None
    for line_index, line in enumerate(lines):
        role = line.role or default_role
        if role == "long_answer" and active_long_answer_event is None:
            active_long_answer_event = event_index
        if role != "long_answer" and active_long_answer_event is not None:
            event_index += 1
            active_long_answer_event = None
        line_is_event = role in ANIMATED_LINE_ROLES
        line_is_target = answer_mode == "overlay" and line_is_event and event_index == target_event_index
        p = ET.SubElement(tx_body, qn(A_NS, "p"))
        p_pr = ET.SubElement(p, qn(A_NS, "pPr"), paragraph_properties(role))
        add_paragraph_line_spacing(p_pr)
        font, size, color, bold = role_style(role)
        line_text = line.text
        number_match = re.match(r"^(\d+[\.、]\s*)", line_text)
        if role in {"body", "reading", "extended_reading", "subheading", "segmentation_question"} and number_match:
            prefix = number_match.group(1)
            append_run(
                p,
                prefix,
                font=font,
                size=size,
                color=QUESTION_NUMBER_COLOR or color,
                bold=True,
                underline=None,
                superscript=False,
                alpha=0 if answer_mode == "overlay" else None,
            )
            line_text = line_text[len(prefix) :]
        for inline in parse_inline_runs(line_text):
            if not inline.text:
                continue
            inline_event_index = active_long_answer_event if role == "long_answer" else event_index
            inline_is_target = answer_mode == "overlay" and inline.answer_fill and inline_event_index == target_event_index
            if answer_mode == "overlay":
                visible = line_is_target or inline_is_target
                alpha = None if visible else 0
                run_underline = inline.underline if visible else None
                run_underline_color = inline.underline_color if visible else None
            elif answer_mode == "base" and (line_is_event or inline.answer_fill):
                alpha = 0
                run_underline = inline.underline
                run_underline_color = inline.underline_color
            else:
                alpha = None
                run_underline = inline.underline
                run_underline_color = inline.underline_color
            run_color = "FF0000" if inline.answer_fill else (inline.color or color)
            run_font = inline.font or ("宋体" if inline.answer_fill else font)
            run_bold = inline.bold if inline.bold is not None else (True if inline.answer_fill else bold)
            run_size = size
            run_italic = role == "title_volume"
            if inline.answer_fill and role != "long_answer" and text_units(strip_inline_tags(inline.text)) > 18.0:
                run_size = max(18.0, size * 0.75)
            # Answer overlays duplicate the base paragraph and hide non-target
            # runs with alpha=0. PowerPoint still paints highlight fills on
            # transparent runs, so strip highlights from hidden overlay text to
            # avoid covering visible highlighted words underneath.
            run_highlight = None if answer_mode == "overlay" and alpha == 0 else inline.highlight
            if inline.emphasis_dot and run_font in {"楷体", "KaiTi"}:
                run_font = "KaiTi"
            if inline.emphasis_dot:
                append_emphasis_dot_runs(
                    p,
                    inline.text,
                    font=run_font,
                    size=run_size,
                    color=run_color,
                    bold=run_bold,
                    underline=run_underline,
                    superscript=inline.superscript,
                    italic=run_italic,
                    underline_color=run_underline_color,
                    highlight=run_highlight,
                    alpha=alpha,
                )
            else:
                append_run(
                    p,
                    inline.text,
                    font=run_font,
                    size=run_size,
                    color=run_color,
                    bold=run_bold,
                    underline=run_underline,
                    superscript=inline.superscript,
                    italic=run_italic,
                    underline_color=run_underline_color,
                    highlight=run_highlight,
                    alpha=alpha,
                )
            if inline.answer_fill and not line_is_event and role != "long_answer":
                event_index += 1
        if line_is_event:
            event_index += 1
        ET.SubElement(p, qn(A_NS, "endParaRPr"), {"lang": "zh-CN", "sz": str(int(size * 100))})

    append_to_sp_tree(root, sp)
    return shape_id


def role_style(role: str) -> tuple[str, float, str, bool]:
    if role == "heading":
        return "黑体", 24, "00AEEF", False
    if role == "subheading":
        return "黑体", 24, "000000", False
    if role == "center_heading":
        return "黑体", 24, "000000", False
    if role == "module":
        return "黑体", 24, "00AEEF", False
    if role == "title_volume":
        return "方正小标宋_GBK", 28, "FFFFFF", False
    if role in {"yilun_table_header", "yilun_table_header_small"}:
        return "宋体", 14, "000000", True
    if role in {"yilun_table_cell", "yilun_table_cell_small"}:
        return "宋体", 14, "000000", False
    if role == "answer_heading":
        return "黑体", 24, "FF0000", True
    if role == "answer":
        return "宋体", 24, "FF0000", True
    if role in {"analysis", "analysis_block"}:
        return "宋体", 24, "FF0000", True
    if role == "segmentation_analysis":
        return "宋体", 24, "FF0000", True
    if role == "segmentation_question":
        return "宋体", 24, "000000", False
    if role == "segmentation_text":
        return "宋体", 24, "000000", False
    if role == "inline_answer_body":
        return "宋体", 24, "000000", False
    if role == "long_answer":
        return "宋体", 24, "000000", False
    if role == "answer_spacer":
        return "宋体", 24, "000000", False
    if role == "reading":
        return "KaiTi", 24, "000000", False
    if role == "extended_reading":
        return "KaiTi", 24, "000000", False
    if role == "source_note":
        return "仿宋", 20, "000000", False
    if role == "annotation":
        return "仿宋", 20, "000000", False
    return "宋体", 24, "000000", False


def count_answer_events(lines: list[StyledLine]) -> int:
    count = 0
    index = 0
    while index < len(lines):
        line = lines[index]
        role = line.role or "body"
        if role in INLINE_FILLED_ANSWER_ROLES:
            index += 1
            continue
        if role == "long_answer":
            has_answer = False
            while index < len(lines) and (lines[index].role or "body") == "long_answer":
                has_answer = has_answer or any(inline.answer_fill for inline in parse_inline_runs(lines[index].text))
                index += 1
            if has_answer:
                count += 1
            continue
        if role in ANIMATED_LINE_ROLES:
            count += 1
            index += 1
            continue
        count += sum(1 for inline in parse_inline_runs(line.text) if inline.answer_fill)
        index += 1
    return count


def append_answer_overlays(
    root: ET.Element,
    *,
    x: float,
    y: float,
    w: float,
    h: float,
    lines: list[StyledLine],
    name: str,
) -> list[AnimationTarget]:
    if any((line.role or "body") in INLINE_FILLED_ANSWER_ROLES for line in lines):
        return []
    animations: list[AnimationTarget] = []
    event_index = 0
    cursor_y = 0.0
    line_index = 0
    overlay_line_factor = 1.32

    def visual_line_count(line: StyledLine) -> int:
        role = line.role or "body"
        if role == "long_answer":
            return 1
        return max(1, estimate_line_count(line.text))

    def overlay_line_step(line: StyledLine) -> float:
        return visual_line_count(line) * line_height_for_role(line.role or "body") * overlay_line_factor

    def event_width(event_lines: list[StyledLine]) -> float:
        text_lines = [strip_inline_tags(line.text).strip() for line in event_lines if strip_inline_tags(line.text).strip()]
        if not text_lines:
            return min(w, 0.35)
        roles = {line.role or "body" for line in event_lines}
        width_pt = estimate_answer_overlay_width_pt(text_lines)
        indent_pad = 0.72 if roles & {"body", "reading", "extended_reading", "long_answer", "analysis", "analysis_block", "segmentation_analysis", "yuedu_analysis"} else 0.0
        return max(0.35, min(w, width_pt / 72.0 + indent_pad + 0.08))

    def event_height(event_lines: list[StyledLine]) -> float:
        total = 0.0
        for event_line in event_lines:
            total += overlay_line_step(event_line)
        return max(0.25, min(h, total + 0.06))

    def append_line_event(event_lines: list[StyledLine], event_y: float, index: int) -> None:
        if not any(strip_inline_tags(line.text).strip() for line in event_lines):
            return
        shape_id = append_textbox(
            root,
            x=x,
            y=event_y,
            w=event_width(event_lines),
            h=event_height(event_lines),
            lines=event_lines,
            name=f"{name} answer {index + 1}",
            answer_mode="normal",
        )
        animations.append(AnimationTarget(shape_id=shape_id, effect="appear", text_target=False))

    while line_index < len(lines):
        line = lines[line_index]
        role = line.role or "body"
        if role in INLINE_FILLED_ANSWER_ROLES:
            cursor_y += overlay_line_step(line)
            line_index += 1
            continue
        if role == "long_answer":
            event_lines: list[StyledLine] = []
            event_y = y + cursor_y
            has_answer = False
            while line_index < len(lines) and (lines[line_index].role or "body") == "long_answer":
                event_line = lines[line_index]
                event_lines.append(event_line)
                has_answer = has_answer or any(inline.answer_fill for inline in parse_inline_runs(event_line.text))
                cursor_y += overlay_line_step(event_line)
                line_index += 1
            if has_answer:
                append_line_event(event_lines, event_y, event_index)
                event_index += 1
            continue

        line_h = overlay_line_step(line)
        if role in ANIMATED_LINE_ROLES:
            append_line_event([line], y + cursor_y, event_index)
            event_index += 1
            cursor_y += line_h
            line_index += 1
            continue

        inline_answer_count = sum(1 for inline in parse_inline_runs(line.text) if inline.answer_fill)
        for _ in range(inline_answer_count):
            shape_id = append_textbox(
                root,
                x=x,
                y=y,
                w=w,
                h=h,
                lines=lines,
                name=f"{name} answer {event_index + 1}",
                answer_mode="overlay",
                target_event_index=event_index,
            )
            animations.append(AnimationTarget(shape_id=shape_id, effect="appear", text_target=False))
            event_index += 1
        cursor_y += line_h
        line_index += 1
    return animations


def line_capacity_for_width(width_in: float) -> float:
    return max(18.0, 31.0 * width_in / 12.25)


def line_height_for_role(role: str) -> float:
    _font, size, _color, _bold = role_style(role)
    return max(0.30, (size / 72.0) * (PARAGRAPH_LINE_SPACING_PCT / 100000))


def collect_tif_positions(
    *,
    x: float,
    y: float,
    w: float,
    lines: list[StyledLine],
) -> list[tuple[str, float, float, float]]:
    positions: list[tuple[str, float, float, float]] = []
    capacity = line_capacity_for_width(w)
    default_char_w = w / capacity
    line_cursor = 0.0

    for line in lines:
        role = line.role or "body"
        line_h = line_height_for_role(role)
        paragraph_indent = 0.70 if role in {"body", "reading", "long_answer"} else 0.0
        inline_units = 0.0
        max_visual_line = 0
        for inline in parse_inline_runs(line.text):
            if inline.tif_name:
                visual_line = int(inline_units // capacity)
                x_offset = (inline_units % capacity) * default_char_w
                if visual_line == 0:
                    x_offset += paragraph_indent
                positions.append(
                    (
                        inline.tif_name,
                        x + x_offset,
                        y + line_cursor + visual_line * line_h + 0.03,
                        min(0.29, line_h * 0.74),
                    )
                )
                inline_units += 1.0
                max_visual_line = max(max_visual_line, visual_line)
                continue
            inline_text = re.sub(r"<[^>]+>", "", html.unescape(inline.text))
            inline_units += text_units(inline_text)
            max_visual_line = max(max_visual_line, int(inline_units // capacity))
        line_cursor += max(1, max_visual_line + 1) * line_h
    return positions


def append_tif_overlays(
    root: ET.Element,
    rels_root: ET.Element,
    media_assets: dict[str, bytes],
    *,
    x: float,
    y: float,
    w: float,
    lines: list[StyledLine],
    tif_search_dirs: list[Path],
) -> None:
    for tif_name, img_x, img_y, img_h in collect_tif_positions(x=x, y=y, w=w, lines=lines):
        tif_path = resolve_tif_path(tif_name, tif_search_dirs)
        if not tif_path:
            continue
        append_tif_picture(
            root,
            rels_root,
            media_assets,
            tif_path=tif_path,
            x=img_x,
            y=img_y,
            h=img_h,
            name=f"fbd tif {tif_name}",
        )


def remove_existing_timing(root: ET.Element) -> None:
    for child in list(root):
        if child.tag == qn(P_NS, "timing"):
            root.remove(child)


def append_start_condition(parent: ET.Element, delay: str) -> None:
    st_cond = ET.SubElement(parent, qn(P_NS, "stCondLst"))
    ET.SubElement(st_cond, qn(P_NS, "cond"), {"delay": delay})


def target_element(shape_id: int, *, text_target: bool) -> ET.Element:
    tgt_el = ET.Element(qn(P_NS, "tgtEl"))
    sp_tgt = ET.SubElement(tgt_el, qn(P_NS, "spTgt"), {"spid": str(shape_id)})
    if text_target:
        tx_el = ET.SubElement(sp_tgt, qn(P_NS, "txEl"))
        ET.SubElement(tx_el, qn(P_NS, "pRg"), {"st": "0", "end": "0"})
    return tgt_el


def append_visibility_set(parent: ET.Element, shape_id: int, *, text_target: bool, next_id) -> None:
    set_el = ET.SubElement(parent, qn(P_NS, "set"))
    c_bhvr = ET.SubElement(set_el, qn(P_NS, "cBhvr"))
    ctn = ET.SubElement(c_bhvr, qn(P_NS, "cTn"), {"id": str(next_id()), "dur": "1", "fill": "hold"})
    append_start_condition(ctn, "0")
    c_bhvr.append(target_element(shape_id, text_target=text_target))
    attr_lst = ET.SubElement(c_bhvr, qn(P_NS, "attrNameLst"))
    ET.SubElement(attr_lst, qn(P_NS, "attrName")).text = "style.visibility"
    to_el = ET.SubElement(set_el, qn(P_NS, "to"))
    ET.SubElement(to_el, qn(P_NS, "strVal"), {"val": "visible"})


def append_appear_effect(parent: ET.Element, target: AnimationTarget, next_id) -> None:
    effect = ET.SubElement(
        parent,
        qn(P_NS, "cTn"),
        {
            "id": str(next_id()),
            "presetID": "1",
            "presetClass": "entr",
            "presetSubtype": "0",
            "fill": "hold",
            "grpId": "0",
            "nodeType": "clickEffect",
        },
    )
    append_start_condition(effect, "0")
    child = ET.SubElement(effect, qn(P_NS, "childTnLst"))
    append_visibility_set(child, target.shape_id, text_target=target.text_target, next_id=next_id)


def append_flyin_effect(parent: ET.Element, target: AnimationTarget, next_id) -> None:
    effect = ET.SubElement(
        parent,
        qn(P_NS, "cTn"),
        {
            "id": str(next_id()),
            "presetID": "2",
            "presetClass": "entr",
            "presetSubtype": "4",
            "fill": "hold",
            "grpId": "0",
            "nodeType": "clickEffect",
        },
    )
    append_start_condition(effect, "0")
    child = ET.SubElement(effect, qn(P_NS, "childTnLst"))
    append_visibility_set(child, target.shape_id, text_target=False, next_id=next_id)

    for attr_name, start_val in (("ppt_x", "#ppt_x"), ("ppt_y", "1+#ppt_h/2")):
        anim = ET.SubElement(child, qn(P_NS, "anim"), {"calcmode": "lin", "valueType": "num"})
        c_bhvr = ET.SubElement(anim, qn(P_NS, "cBhvr"), {"additive": "base"})
        ET.SubElement(c_bhvr, qn(P_NS, "cTn"), {"id": str(next_id()), "dur": "500", "fill": "hold"})
        c_bhvr.append(target_element(target.shape_id, text_target=False))
        attr_lst = ET.SubElement(c_bhvr, qn(P_NS, "attrNameLst"))
        ET.SubElement(attr_lst, qn(P_NS, "attrName")).text = attr_name
        tav_lst = ET.SubElement(anim, qn(P_NS, "tavLst"))
        tav0 = ET.SubElement(tav_lst, qn(P_NS, "tav"), {"tm": "0"})
        val0 = ET.SubElement(tav0, qn(P_NS, "val"))
        ET.SubElement(val0, qn(P_NS, "strVal"), {"val": start_val})
        tav1 = ET.SubElement(tav_lst, qn(P_NS, "tav"), {"tm": "100000"})
        val1 = ET.SubElement(tav1, qn(P_NS, "val"))
        ET.SubElement(val1, qn(P_NS, "strVal"), {"val": f"#{attr_name}"})


def append_click_effect(main_child_list: ET.Element, target: AnimationTarget, next_id) -> None:
    outer_par = ET.SubElement(main_child_list, qn(P_NS, "par"))
    outer_ctn = ET.SubElement(outer_par, qn(P_NS, "cTn"), {"id": str(next_id()), "fill": "hold"})
    append_start_condition(outer_ctn, "indefinite")
    outer_child = ET.SubElement(outer_ctn, qn(P_NS, "childTnLst"))
    mid_par = ET.SubElement(outer_child, qn(P_NS, "par"))
    mid_ctn = ET.SubElement(mid_par, qn(P_NS, "cTn"), {"id": str(next_id()), "fill": "hold"})
    append_start_condition(mid_ctn, "0")
    mid_child = ET.SubElement(mid_ctn, qn(P_NS, "childTnLst"))
    effect_par = ET.SubElement(mid_child, qn(P_NS, "par"))
    if target.effect == "flyin":
        append_flyin_effect(effect_par, target, next_id)
    else:
        append_appear_effect(effect_par, target, next_id)


def set_slide_animations(root: ET.Element, animations: list[AnimationTarget]) -> None:
    remove_existing_timing(root)
    return

    id_value = 1

    def next_id() -> int:
        nonlocal id_value
        current = id_value
        id_value += 1
        return current

    timing = ET.Element(qn(P_NS, "timing"))
    tn_lst = ET.SubElement(timing, qn(P_NS, "tnLst"))
    root_par = ET.SubElement(tn_lst, qn(P_NS, "par"))
    root_ctn = ET.SubElement(
        root_par,
        qn(P_NS, "cTn"),
        {"id": str(next_id()), "dur": "indefinite", "restart": "never", "nodeType": "tmRoot"},
    )
    root_child = ET.SubElement(root_ctn, qn(P_NS, "childTnLst"))
    seq = ET.SubElement(root_child, qn(P_NS, "seq"), {"concurrent": "1", "nextAc": "seek"})
    main_ctn = ET.SubElement(seq, qn(P_NS, "cTn"), {"id": str(next_id()), "dur": "indefinite", "nodeType": "mainSeq"})
    main_child = ET.SubElement(main_ctn, qn(P_NS, "childTnLst"))

    for target in animations:
        append_click_effect(main_child, target, next_id)

    prev_cond = ET.SubElement(seq, qn(P_NS, "prevCondLst"))
    cond = ET.SubElement(prev_cond, qn(P_NS, "cond"), {"evt": "onPrev", "delay": "0"})
    tgt = ET.SubElement(cond, qn(P_NS, "tgtEl"))
    ET.SubElement(tgt, qn(P_NS, "sldTgt"))
    next_cond = ET.SubElement(seq, qn(P_NS, "nextCondLst"))
    cond = ET.SubElement(next_cond, qn(P_NS, "cond"), {"evt": "onNext", "delay": "0"})
    tgt = ET.SubElement(cond, qn(P_NS, "tgtEl"))
    ET.SubElement(tgt, qn(P_NS, "sldTgt"))

    bld_lst = ET.SubElement(timing, qn(P_NS, "bldLst"))
    for target in animations:
        attrs = {"spid": str(target.shape_id), "grpId": "0"}
        if target.text_target:
            attrs["build"] = "allAtOnce"
        else:
            attrs["animBg"] = "1"
        ET.SubElement(bld_lst, qn(P_NS, "bldP"), attrs)

    ext_lst = root.find(qn(P_NS, "extLst"))
    if ext_lst is not None:
        root.insert(list(root).index(ext_lst), timing)
    else:
        root.append(timing)


def add_cover_content(root: ET.Element, volume: str, start: int, end: int) -> None:
    # Reuse the template's own cover text boxes so the title size/position stay intact.
    replace_exact_text(root, "必修上册", volume)
    piece_range = f"篇目{start}-{end}"
    replace_exact_text(root, "篇目1-4", piece_range)
    for sp in text_shapes(root):
        text = get_shape_text(sp).strip()
        if re.fullmatch(r"篇目\d+-\d+", text):
            set_shape_text(sp, piece_range)
            text = piece_range
        if text in {volume, piece_range, f"{volume} {piece_range}", f"{volume}{piece_range}"}:
            style_existing_text_shape(sp, font="宋体", color="000000", bold=False)


def directory_item_y_positions(count: int) -> list[float]:
    if count <= 3:
        return [1.897, 3.069, 4.242][:count]
    if count == 4:
        return [1.50, 2.57, 3.64, 4.71]
    start = 1.28
    step = min(0.92, 4.55 / max(1, count - 1))
    return [start + i * step for i in range(count)]


def append_directory_item(
    root: ET.Element,
    rels_root: ET.Element,
    *,
    piece: Piece,
    slide_number: int,
    x: float,
    y: float,
) -> int:
    shape_id = next_shape_id(root)
    sp = ET.Element(qn(P_NS, "sp"))

    nv_sp_pr = ET.SubElement(sp, qn(P_NS, "nvSpPr"))
    c_nv_pr = ET.SubElement(nv_sp_pr, qn(P_NS, "cNvPr"), {"id": str(shape_id), "name": f"目录项 {piece.num}"})
    rel_id = add_slide_jump_relationship(rels_root, slide_number)
    ET.SubElement(c_nv_pr, qn(A_NS, "hlinkClick"), {qn(R_NS, "id"): rel_id, "action": "ppaction://hlinksldjump"})
    ET.SubElement(nv_sp_pr, qn(P_NS, "cNvSpPr"))
    ET.SubElement(nv_sp_pr, qn(P_NS, "nvPr"))

    sp_pr = ET.SubElement(sp, qn(P_NS, "spPr"))
    xfrm = ET.SubElement(sp_pr, qn(A_NS, "xfrm"))
    ET.SubElement(xfrm, qn(A_NS, "off"), {"x": str(inch(x)), "y": str(inch(y))})
    ET.SubElement(xfrm, qn(A_NS, "ext"), {"cx": str(inch(5.65)), "cy": str(inch(0.91))})
    geom = ET.SubElement(sp_pr, qn(A_NS, "prstGeom"), {"prst": "rect"})
    ET.SubElement(geom, qn(A_NS, "avLst"))
    ET.SubElement(sp_pr, qn(A_NS, "noFill"))
    ln = ET.SubElement(sp_pr, qn(A_NS, "ln"), {"w": "9525"})
    ET.SubElement(ln, qn(A_NS, "noFill"))

    tx_body = ET.SubElement(sp, qn(P_NS, "txBody"))
    body_pr = ET.SubElement(tx_body, qn(A_NS, "bodyPr"), {"wrap": "square", "anchor": "t"})
    ET.SubElement(body_pr, qn(A_NS, "spAutoFit"))
    ET.SubElement(tx_body, qn(A_NS, "lstStyle"))
    p = ET.SubElement(tx_body, qn(A_NS, "p"))
    p_pr = ET.SubElement(p, qn(A_NS, "pPr"), {"algn": "l"})
    ln_spc = ET.SubElement(p_pr, qn(A_NS, "lnSpc"))
    ET.SubElement(ln_spc, qn(A_NS, "spcPct"), {"val": "150000"})
    append_run(p, f"{piece.num}. ", font="黑体", size=32, color="000000", bold=True)
    append_run(p, piece.title, font="黑体", size=32, color="000000", bold=True)
    ET.SubElement(p, qn(A_NS, "endParaRPr"), {"lang": "zh-CN", "sz": "3200"})
    append_to_sp_tree(root, sp)
    return shape_id


def add_directory_content(root: ET.Element, rels_root: ET.Element, pieces: list[Piece], title_slide_numbers: dict[int, int]) -> None:
    for sp in text_shapes(root):
        if get_shape_text(sp).strip() and "各篇目点击即可直达" not in get_shape_text(sp):
            set_shape_text(sp, "")

    item_x = 5.95
    line_x = 5.20
    line_w = 4.52
    for piece, y in zip(pieces, directory_item_y_positions(len(pieces))):
        append_directory_item(
            root,
            rels_root,
            piece=piece,
            slide_number=title_slide_numbers[piece.num],
            x=item_x,
            y=y,
        )
        append_dashed_line(root, x=line_x, y=y + 0.77, w=line_w, name=f"目录虚线 {piece.num}")


def piece_title_font_size(title: str) -> float:
    units = text_units(title)
    if "★" in title:
        return 72
    if units <= 6:
        return 80
    if units <= 9:
        return 72
    if units <= 12:
        return 62
    return 54


def split_title_suffix(title: str) -> tuple[str, str]:
    for suffix in ("并序", "序"):
        if title.endswith(suffix) and len(title) > len(suffix):
            return title[: -len(suffix)], suffix
    return title, ""


def append_piece_title_textbox(root: ET.Element, piece: Piece) -> None:
    size = piece_title_font_size(piece.title)
    main_title, suffix = split_title_suffix(piece.title)
    shape_id = next_shape_id(root)
    sp = ET.Element(qn(P_NS, "sp"))

    nv_sp_pr = ET.SubElement(sp, qn(P_NS, "nvSpPr"))
    ET.SubElement(nv_sp_pr, qn(P_NS, "cNvPr"), {"id": str(shape_id), "name": f"篇目标题 {piece.num}"})
    ET.SubElement(nv_sp_pr, qn(P_NS, "cNvSpPr"), {"txBox": "1"})
    ET.SubElement(nv_sp_pr, qn(P_NS, "nvPr"))

    sp_pr = ET.SubElement(sp, qn(P_NS, "spPr"))
    xfrm = ET.SubElement(sp_pr, qn(A_NS, "xfrm"))
    ET.SubElement(xfrm, qn(A_NS, "off"), {"x": str(inch(1.75)), "y": str(inch(2.37))})
    ET.SubElement(xfrm, qn(A_NS, "ext"), {"cx": str(inch(9.85)), "cy": str(inch(1.55))})
    geom = ET.SubElement(sp_pr, qn(A_NS, "prstGeom"), {"prst": "rect"})
    ET.SubElement(geom, qn(A_NS, "avLst"))
    ET.SubElement(sp_pr, qn(A_NS, "noFill"))
    ln = ET.SubElement(sp_pr, qn(A_NS, "ln"), {"w": "9525"})
    ET.SubElement(ln, qn(A_NS, "noFill"))

    tx_body = ET.SubElement(sp, qn(P_NS, "txBody"))
    body_pr = ET.SubElement(tx_body, qn(A_NS, "bodyPr"), {"wrap": "square", "anchor": "t"})
    ET.SubElement(body_pr, qn(A_NS, "spAutoFit"))
    ET.SubElement(tx_body, qn(A_NS, "lstStyle"))
    p = ET.SubElement(tx_body, qn(A_NS, "p"))
    ET.SubElement(p, qn(A_NS, "pPr"), {"algn": "ctr"})
    append_run(p, f"{piece.num}. ", font="方正小标宋_GBK", size=size, color="BD1F6C", bold=True)
    append_run(p, main_title, font="方正小标宋_GBK", size=size, color="000000", bold=True)
    if suffix:
        append_run(p, suffix, font="楷体", size=size, color="000000", bold=False)
    ET.SubElement(p, qn(A_NS, "endParaRPr"), {"lang": "zh-CN", "sz": str(int(size * 100))})
    append_to_sp_tree(root, sp)


def spaced_volume_label(volume: str) -> str:
    return " ".join(volume)


def append_title_volume_label(root: ET.Element, volume: str) -> None:
    append_textbox(
        root,
        x=0.38,
        y=6.42,
        w=2.55,
        h=0.52,
        lines=[StyledLine(spaced_volume_label(volume), "title_volume")],
        name="篇目标题册别",
    )


def add_piece_title(root: ET.Element, piece: Piece) -> None:
    clear_text_shapes(root)
    append_piece_title_textbox(root, piece)
    append_title_volume_label(root, piece.volume)


def split_module_header(chunk: list[StyledLine]) -> tuple[str | None, list[StyledLine]]:
    if chunk and chunk[0].role == "module_header":
        return canonical_module_title(chunk[0].text), chunk[1:]
    return None, chunk


def add_segmentation_content(
    root: ET.Element,
    rels_root: ET.Element,
    piece: Piece,
    chunk: list[StyledLine],
    module_title: str | None = None,
) -> None:
    clear_text_shapes(root, keep_substrings=("能力",))
    if module_title:
        add_module_header_image(root, rels_root, module_title)
    question = next((line for line in chunk if line.role == "segmentation_question"), None)
    sentence = next((line for line in chunk if line.role == "segmentation_text"), None)
    analyses = [line for line in chunk if line.role in {"segmentation_analysis", "yuedu_analysis"}]
    if question is None or sentence is None:
        y = 1.12 if module_title else 0.55
        h = 5.25 if module_title else 5.85
        append_textbox(root, x=0.55, y=y, w=12.25, h=h, lines=chunk, name=f"{piece.title} content")
        set_slide_animations(root, [])
        return

    x = 0.58
    y_shift = 0.20 if module_title else 0.0
    question_y = 1.34 + y_shift
    sentence_w = 12.20
    question_lines = max(1, math.ceil(text_units(strip_inline_tags(question.text)) / line_capacity_for_width(sentence_w)))
    question_h = max(0.45, question_lines * line_height_for_role("segmentation_question"))
    sentence_y = question_y + question_h + 0.12
    sentence_h = 0.88 if segmentation_text_width(sentence.text) > sentence_w else 0.45
    analysis_y = sentence_y + sentence_h + 0.26
    box_y = sentence_y + 0.02
    box_w = 0.33
    box_h = 0.44
    answer = str((sentence.meta or {}).get("answer", ""))
    boxes = segmentation_letter_boxes(sentence.text, x)
    animations: list[AnimationTarget] = []

    for letter, box_x in boxes:
        if letter in answer:
            shape_id = append_rect(
                root,
                x=box_x,
                y=box_y,
                w=box_w,
                h=box_h,
                fill_color="5B9BD5",
                line_color=None,
                name="segmentation answer fill",
            )
            animations.append(AnimationTarget(shape_id=shape_id, effect="flyin", text_target=False))

    append_textbox(
        root,
        x=x,
        y=question_y,
        w=12.20,
        h=question_h,
        lines=[question],
        name=f"{piece.title} segmentation question",
    )
    append_textbox(
        root,
        x=x,
        y=sentence_y,
        w=sentence_w,
        h=sentence_h,
        lines=[sentence],
        name=f"{piece.title} segmentation text",
    )
    if analyses:
        analysis_shape_id = append_textbox(
            root,
            x=x,
            y=analysis_y,
            w=11.8,
            h=2.3,
            lines=analyses,
            name=f"{piece.title} segmentation analysis",
        )
        animations.append(AnimationTarget(shape_id=analysis_shape_id, effect="appear", text_target=False))

    for _letter, box_x in boxes:
        append_rect(
            root,
            x=box_x,
            y=box_y,
            w=box_w,
            h=box_h,
            fill_color=None,
            line_color="000000",
            name="segmentation option box",
        )
    set_slide_animations(root, animations)


def add_content(
    root: ET.Element,
    rels_root: ET.Element,
    piece: Piece,
    chunk: list[StyledLine],
    *,
    media_assets: dict[str, bytes],
    tif_search_dirs: list[Path],
    content_box: tuple[float, float, float, float],
) -> None:
    clear_text_shapes(root, keep_substrings=("能力",))
    module_title, content_lines = split_module_header(chunk)
    if module_title:
        add_module_header_image(root, rels_root, module_title)
    if any(line.role == "segmentation_text" for line in content_lines):
        add_segmentation_content(root, rels_root, piece, content_lines, module_title=None)
        return
    base_x, base_y, base_w, base_h = content_box
    y = max(base_y, 1.12) if module_title else base_y
    h = base_h - (y - base_y)
    append_textbox(
        root,
        x=base_x,
        y=y,
        w=base_w,
        h=h,
        lines=content_lines,
        name=f"{piece.title} content",
        answer_mode="base",
    )
    append_tif_overlays(
        root,
        rels_root,
        media_assets,
        x=base_x,
        y=y,
        w=base_w,
        lines=content_lines,
        tif_search_dirs=tif_search_dirs,
    )
    animations = append_answer_overlays(
        root,
        x=base_x,
        y=y,
        w=base_w,
        h=h,
        lines=content_lines,
        name=f"{piece.title} content",
    )
    set_slide_animations(root, animations)


def parse_xml(data: bytes) -> ET.Element:
    return ET.fromstring(data)


def update_content_types(root: ET.Element, slide_count: int) -> bytes:
    for child in list(root):
        if local_name(child.tag) == "Override" and child.attrib.get("PartName", "").startswith("/ppt/slides/slide"):
            root.remove(child)
    for i in range(1, slide_count + 1):
        ET.SubElement(
            root,
            qn(CT_NS, "Override"),
            {
                "PartName": f"/ppt/slides/slide{i}.xml",
                "ContentType": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
            },
        )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def update_presentation_xml(root: ET.Element, slide_count: int) -> bytes:
    sld_id_lst = root.find(qn(P_NS, "sldIdLst"))
    if sld_id_lst is None:
        raise ValueError("presentation.xml has no sldIdLst")
    for child in list(sld_id_lst):
        sld_id_lst.remove(child)
    for i in range(1, slide_count + 1):
        ET.SubElement(
            sld_id_lst,
            qn(P_NS, "sldId"),
            {"id": str(255 + i), qn(R_NS, "id"): f"rIdSlide{i}"},
        )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def update_presentation_rels(root: ET.Element, slide_count: int) -> bytes:
    for child in list(root):
        target = child.attrib.get("Target", "")
        rel_type = child.attrib.get("Type", "")
        if rel_type.endswith("/slide") or target.startswith("slides/slide"):
            root.remove(child)
    for i in range(1, slide_count + 1):
        ET.SubElement(
            root,
            qn(REL_NS, "Relationship"),
            {
                "Id": f"rIdSlide{i}",
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                "Target": f"slides/slide{i}.xml",
            },
        )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def copy_slide_rels_root(zf: zipfile.ZipFile, src_index: int) -> ET.Element:
    name = f"ppt/slides/_rels/slide{src_index}.xml.rels"
    try:
        root = parse_xml(zf.read(name))
    except KeyError:
        return ET.Element(qn(REL_NS, "Relationships"))
    for rel in list(root):
        if rel.attrib.get("Type", "").endswith("/tags"):
            root.remove(rel)
    return root


def make_slide_parts(zf: zipfile.ZipFile, src_index: int, edit_func) -> tuple[bytes, bytes]:
    root = parse_xml(zf.read(f"ppt/slides/slide{src_index}.xml"))
    rels_root = copy_slide_rels_root(zf, src_index)
    strip_copied_custom_data(root)
    edit_func(root, rels_root)
    return (
        ET.tostring(root, encoding="utf-8", xml_declaration=True),
        ET.tostring(rels_root, encoding="utf-8", xml_declaration=True),
    )


def template_contract(slide_count: int) -> tuple[int, int, list[int], tuple[float, float, float, float]]:
    if slide_count <= 6:
        return 4, 3, [5, 6], (2.45, 0.42, 10.35, 6.25)
    return 4, 5, [slide_count - 2, slide_count - 1, slide_count], (0.55, 0.55, 12.25, 5.85)


def run_has_zero_alpha(run: ET.Element) -> bool:
    r_pr = run.find(qn(A_NS, "rPr"))
    if r_pr is None:
        return False
    for alpha in r_pr.findall(f".//{{{A_NS}}}alpha"):
        if alpha.attrib.get("val") == "0":
            return True
    return False


def trim_range_text(start: int, text: str) -> tuple[int, str]:
    left_trimmed = text.lstrip(" \t\r\n　\u00a0_＿")
    leading = len(text) - len(left_trimmed)
    trimmed = left_trimmed.rstrip(" \t\r\n　\u00a0_＿")
    return start + leading, trimmed


def estimate_answer_overlay_width_pt(text_lines: list[str]) -> float:
    units = max((text_units(line.strip()) for line in text_lines if line.strip()), default=0.0)
    return max(18.0, min(840.0, units * 24.0 + 8.0))


def should_compact_answer_overlay(text_lines: list[str]) -> bool:
    text = "".join(line.strip() for line in text_lines)
    if not text:
        return False
    if re.match(r"^【(?:解析|关键能力|解题思路|解决思路)】", text):
        return False
    units = max((text_units(line.strip()) for line in text_lines if line.strip()), default=0.0)
    return len([line for line in text_lines if line.strip()]) == 1 and units <= 18.0


def collect_answer_overlay_specs(pptx: Path) -> dict[tuple[int, str], AnswerOverlaySpec]:
    specs: dict[tuple[int, str], AnswerOverlaySpec] = {}
    with zipfile.ZipFile(pptx, "r") as zf:
        for name in zf.namelist():
            match = re.match(r"ppt/slides/slide(\d+)\.xml$", name)
            if not match:
                continue
            slide_index = int(match.group(1))
            root = parse_xml(zf.read(name))
            for sp in root.findall(f".//{{{P_NS}}}sp"):
                c_nv_pr = sp.find(f".//{{{P_NS}}}cNvPr")
                if c_nv_pr is None:
                    continue
                shape_name = c_nv_pr.attrib.get("name", "")
                if " answer " not in shape_name:
                    continue
                tx_body = sp.find(qn(P_NS, "txBody"))
                if tx_body is None:
                    continue

                char_index = 1
                visible_ranges: list[tuple[int, int]] = []
                visible_lines: list[str] = []
                inline_context = False
                for paragraph in tx_body.findall(qn(A_NS, "p")):
                    paragraph_visible: list[str] = []
                    paragraph_ranges: list[tuple[int, int]] = []
                    paragraph_hidden_text = False
                    for run in paragraph.findall(qn(A_NS, "r")):
                        text_el = run.find(qn(A_NS, "t"))
                        text = text_el.text if text_el is not None and text_el.text is not None else ""
                        if text and not run_has_zero_alpha(run):
                            range_start, trimmed = trim_range_text(char_index, text)
                            if trimmed:
                                paragraph_visible.append(trimmed)
                                paragraph_ranges.append((range_start, len(trimmed)))
                        elif text and run_has_zero_alpha(run) and text.strip():
                            paragraph_hidden_text = True
                        char_index += len(text)
                    if paragraph_visible:
                        if paragraph_hidden_text:
                            inline_context = True
                        visible_lines.append("".join(paragraph_visible).strip())
                        visible_ranges.extend(paragraph_ranges)
                    char_index += 1
                if visible_ranges and visible_lines:
                    specs[(slide_index, shape_name)] = AnswerOverlaySpec(
                        visible_ranges,
                        visible_lines,
                        estimate_answer_overlay_width_pt(visible_lines),
                        should_compact_answer_overlay(visible_lines),
                        inline_context,
                    )
    return specs


def answer_animation_index(shape_name: str) -> int | None:
    match = re.search(r"\banswer\s+(\d+)\b", shape_name, flags=re.I)
    return int(match.group(1)) if match else None


def clear_powerpoint_animation_sequence(slide) -> None:
    try:
        sequence = slide.TimeLine.MainSequence
        for index in range(sequence.Count, 0, -1):
            sequence.Item(index).Delete()
    except Exception:
        pass


def add_powerpoint_standard_animations(slide) -> int:
    answer_shapes: list[tuple[int, object]] = []
    segmentation_fills: list[object] = []
    segmentation_analyses: list[object] = []
    for shape in slide.Shapes:
        try:
            name = shape.Name
            lower_name = name.lower()
        except Exception:
            continue
        if "segmentation answer fill" in lower_name:
            segmentation_fills.append(shape)
            continue
        if "segmentation analysis" in lower_name:
            segmentation_analyses.append(shape)
            continue
        order = answer_animation_index(name)
        if order is not None:
            answer_shapes.append((order, shape))

    if not answer_shapes and not segmentation_fills and not segmentation_analyses:
        return 0

    clear_powerpoint_animation_sequence(slide)
    sequence = slide.TimeLine.MainSequence
    count = 0
    grouped_answer_shapes: dict[int, list[object]] = {}
    for order, shape in sorted(answer_shapes, key=lambda item: (item[0], float(item[1].Top), float(item[1].Left))):
        grouped_answer_shapes.setdefault(order, []).append(shape)
    for _order, shapes in grouped_answer_shapes.items():
        for shape_index, shape in enumerate(shapes):
            try:
                trigger = 1 if shape_index == 0 else 2  # on click, then with previous for split fragments
                sequence.AddEffect(shape, 1, 0, trigger)  # appear
                count += 1
            except Exception:
                pass
    for shape in sorted(segmentation_fills, key=lambda item: (float(item.Top), float(item.Left))):
        try:
            sequence.AddEffect(shape, 2, 0, 1)  # fly in, on click
            count += 1
        except Exception:
            pass
    for shape in sorted(segmentation_analyses, key=lambda item: (float(item.Top), float(item.Left))):
        try:
            sequence.AddEffect(shape, 1, 0, 1)
            count += 1
        except Exception:
            pass
    return count


def style_cover_subtitle_with_powerpoint(pres) -> int:
    try:
        slide = pres.Slides(1)
    except Exception:
        return 0
    count = 0
    for shape in slide.Shapes:
        try:
            if not (shape.HasTextFrame and shape.TextFrame.HasText):
                continue
            text = shape.TextFrame.TextRange.Text.strip()
        except Exception:
            continue
        if not (text.startswith("篇目") or text in VOLUMES):
            continue
        try:
            font = shape.TextFrame.TextRange.Font
            font.Name = "宋体"
            font.NameFarEast = "宋体"
            font.Bold = False
            font.Color.RGB = 0
            count += 1
        except Exception:
            continue
    return count


def clear_answer_overlay_alpha(pptx: Path) -> None:
    rewritten: dict[str, bytes] = {}
    changed = 0
    with zipfile.ZipFile(pptx, "r") as zin:
        for name in zin.namelist():
            if not re.match(r"ppt/slides/slide\d+\.xml$", name):
                continue
            root = parse_xml(zin.read(name))
            slide_changed = False
            for sp in root.findall(f".//{{{P_NS}}}sp"):
                c_nv_pr = sp.find(f".//{{{P_NS}}}cNvPr")
                shape_name = c_nv_pr.attrib.get("name", "") if c_nv_pr is not None else ""
                lower_name = shape_name.lower()
                if "answer-adjusted" not in lower_name:
                    continue
                for alpha in list(sp.findall(f".//{{{A_NS}}}alpha")):
                    if remove_descendant(sp, alpha):
                        slide_changed = True
                        changed += 1
            if slide_changed:
                rewritten[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

        if not rewritten:
            return

        temp_zip = pptx.with_suffix(".alpha_fix.tmp")
        with zipfile.ZipFile(temp_zip, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.is_dir():
                    continue
                data = rewritten.get(item.filename)
                if data is None:
                    data = zin.read(item.filename)
                zout.writestr(item, data)
    shutil.move(str(temp_zip), str(pptx))
    print(f"答案透明属性清理：{pptx.name}，已清理 {changed} 处")


def powerpoint_char_boxes_for_text(shape, text_lines: list[str]) -> list[tuple[str, float, float, float, float]]:
    char_boxes: list[tuple[str, float, float, float, float]] = []
    try:
        text_range = shape.TextFrame.TextRange
        full_text = text_range.Text
    except Exception:
        return char_boxes
    search_from = 0
    for line in text_lines:
        needle = line.strip()
        if not needle:
            continue
        found = full_text.find(needle, search_from)
        if found < 0:
            found = full_text.find(needle)
        if found < 0:
            continue
        for char_index in range(found + 1, found + 1 + len(needle)):
            try:
                char_range = text_range.Characters(char_index, 1)
                char_boxes.append(
                    (
                        char_range.Text,
                        float(char_range.BoundLeft),
                        float(char_range.BoundTop),
                        float(char_range.BoundLeft + char_range.BoundWidth),
                        float(char_range.BoundTop + char_range.BoundHeight),
                    )
                )
            except Exception:
                continue
        search_from = found + len(needle)
    return char_boxes


def powerpoint_char_boxes_for_ranges(shape, ranges: list[tuple[int, int]]) -> list[tuple[str, float, float, float, float]]:
    char_boxes: list[tuple[str, float, float, float, float]] = []
    for start, length in ranges:
        for char_index in range(start, start + length):
            try:
                text_range = shape.TextFrame.TextRange.Characters(char_index, 1)
                char_text = text_range.Text
                char_boxes.append(
                    (
                        char_text,
                        float(text_range.BoundLeft),
                        float(text_range.BoundTop),
                        float(text_range.BoundLeft + text_range.BoundWidth),
                        float(text_range.BoundTop + text_range.BoundHeight),
                    )
                )
            except Exception:
                continue
    return char_boxes


def answer_overlay_base_name(shape_name: str) -> str | None:
    match = re.search(r"\s+answer\s+\d+\b", shape_name, flags=re.I)
    if not match:
        return None
    return shape_name[: match.start()].strip()


def find_answer_base_shape(slide, shape_name: str):
    base_name = answer_overlay_base_name(shape_name)
    if not base_name:
        return None
    base_lower = base_name.lower()
    for candidate in slide.Shapes:
        try:
            candidate_name = candidate.Name
            lower_name = candidate_name.lower()
            if "answer" in lower_name:
                continue
            if not lower_name.startswith(base_lower):
                continue
            if not (candidate.HasTextFrame and candidate.TextFrame.HasText):
                continue
            return candidate
        except Exception:
            continue
    return None


def should_anchor_answer_to_base_shape(shape_name: str) -> bool:
    lower_name = shape_name.lower()
    return " open line " in lower_name


def dominant_visual_line_char_boxes(
    char_boxes: list[tuple[str, float, float, float, float]]
) -> list[tuple[str, float, float, float, float]]:
    if not char_boxes:
        return []
    groups: list[list[tuple[str, float, float, float, float]]] = []
    for box in char_boxes:
        for group in groups:
            if abs(box[2] - group[0][2]) <= 5:
                group.append(box)
                break
        else:
            groups.append([box])
    groups.sort(key=lambda group: (-len(group), group[0][2], min(box[1] for box in group)))
    return sorted(groups[0], key=lambda box: box[1])


def exported_shape_visual_bounds(shape, export_path: Path) -> tuple[float, float, float, float] | None:
    try:
        from PIL import Image
    except Exception:
        return None
    image = None
    try:
        shape.Export(str(export_path), 2)
        image = Image.open(export_path).convert("RGBA")
    except Exception:
        return None
    try:
        width_px, height_px = image.size
        if width_px <= 0 or height_px <= 0:
            return None
        pixels = image.load()
        xs: list[int] = []
        ys: list[int] = []
        for py in range(height_px):
            for px in range(width_px):
                red, green, blue, alpha = pixels[px, py]
                if alpha <= 10:
                    continue
                if red > 248 and green > 248 and blue > 248:
                    continue
                xs.append(px)
                ys.append(py)
        if not xs:
            return None
        scale_x = float(shape.Width) / width_px
        scale_y = float(shape.Height) / height_px
        return (
            float(shape.Left) + min(xs) * scale_x,
            float(shape.Top) + min(ys) * scale_y,
            float(shape.Left) + (max(xs) + 1) * scale_x,
            float(shape.Top) + (max(ys) + 1) * scale_y,
        )
    finally:
        try:
            if image is not None:
                image.close()
        except Exception:
            pass


def align_shape_text_to_bounds(shape, desired_left: float, desired_top: float) -> None:
    for _ in range(3):
        try:
            text_range = shape.TextFrame.TextRange
            bound_left = float(text_range.BoundLeft)
            bound_top = float(text_range.BoundTop)
        except Exception:
            return
        dx = desired_left - bound_left
        dy = desired_top - bound_top
        if abs(dx) < 0.5 and abs(dy) < 0.5:
            return
        try:
            shape.Left = float(shape.Left) + dx
            shape.Top = float(shape.Top) + dy
        except Exception:
            return


def crop_answer_overlays_with_powershell(pptx: Path) -> bool:
    """Crop answer overlay shapes using PowerPoint COM when win32com is absent."""

    if sys.platform != "win32":
        return False
    specs = collect_answer_overlay_specs(pptx)
    if not specs:
        return False
    powershell = shutil.which("powershell") or shutil.which("pwsh")
    if not powershell:
        return False
    spec_path = pptx.with_name(f"_{pptx.stem}_answer_overlay_specs.json")
    spec_payload = [
        {
            "slide": slide_index,
            "name": shape_name,
            "ranges": [list(item) for item in spec.ranges],
            "text": "\r".join(line for line in spec.text_lines if line),
            "inline": spec.inline_context,
            "compact": spec.compact,
            "expectedWidth": spec.expected_width_pt,
            "analysis": any(
                label in "".join(spec.text_lines)
                for label in ("【解析】", "【关键能力】", "【解题思路】", "【解决思路】")
            ),
        }
        for (slide_index, shape_name), spec in specs.items()
    ]
    spec_path.write_text(json.dumps(spec_payload, ensure_ascii=False), encoding="utf-8")
    script = r'''
$ErrorActionPreference = "Stop"
$pptx = $env:PPTX_TO_CROP
if (-not $pptx) { throw "PPTX_TO_CROP is empty." }
$pptx = (Resolve-Path -LiteralPath $pptx).Path
$specPath = $env:ANSWER_SPEC_JSON
if (-not $specPath) { throw "ANSWER_SPEC_JSON is empty." }
$specItems = Get-Content -Raw -Encoding UTF8 -LiteralPath $specPath | ConvertFrom-Json
$specMap = @{}
foreach ($item in @($specItems)) {
    $specMap["$($item.slide)|$($item.name)"] = $item
}
function Get-AnswerBaseShape($slide, [string]$shapeName) {
    $match = [regex]::Match($shapeName, "\s+answer\s+\d+\b")
    if (-not $match.Success) { return $null }
    $baseName = $shapeName.Substring(0, $match.Index).Trim().ToLowerInvariant()
    if (-not $baseName) { return $null }
    for ($candidateIndex = 1; $candidateIndex -le $slide.Shapes.Count; $candidateIndex++) {
        try {
            $candidate = $slide.Shapes.Item($candidateIndex)
            $candidateName = ([string]$candidate.Name).ToLowerInvariant()
            if ($candidateName.Contains("answer")) { continue }
            if (-not $candidateName.StartsWith($baseName)) { continue }
            if ($candidate.HasTextFrame -eq 0) { continue }
            if ($candidate.TextFrame.HasText -eq 0) { continue }
            return $candidate
        } catch {
            continue
        }
    }
    return $null
}
function Get-RangeBoxes($shape, $ranges) {
    $boxes = @()
    foreach ($range in @($ranges)) {
        try {
            $start = [int]$range[0]
            $length = [int]$range[1]
            if ($start -le 0 -or $length -le 0) { continue }
            $charRange = $shape.TextFrame.TextRange.Characters($start, $length)
            $boxes += [PSCustomObject]@{
                Left = [double]$charRange.BoundLeft
                Top = [double]$charRange.BoundTop
                Right = [double]$charRange.BoundLeft + [double]$charRange.BoundWidth
                Bottom = [double]$charRange.BoundTop + [double]$charRange.BoundHeight
            }
        } catch {
            continue
        }
    }
    return $boxes
}
$pp = $null
$pres = $null
$adjusted = 0
try {
    $pp = New-Object -ComObject PowerPoint.Application
    $pres = $pp.Presentations.Open($pptx, $false, $false, $false)
    for ($si = 1; $si -le $pres.Slides.Count; $si++) {
        $slide = $pres.Slides.Item($si)
        for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
            $shape = $slide.Shapes.Item($shapeIndex)
            try {
                $shapeName = [string]$shape.Name
                if ($shapeName -notmatch "\sanswer\s\d+\b") { continue }
                $specKey = "$si|$shapeName"
                if (-not $specMap.ContainsKey($specKey)) { continue }
                $spec = $specMap[$specKey]
                if ($shape.HasTextFrame -eq 0) { continue }
                if ($shape.TextFrame.HasText -eq 0) { continue }

                $sourceShape = $shape
                if ([bool]$spec.inline) {
                    $baseShape = Get-AnswerBaseShape $slide $shapeName
                    if ($baseShape -eq $null) { continue }
                    $sourceShape = $baseShape
                }
                $boxes = @(Get-RangeBoxes $sourceShape $spec.ranges)
                if ($boxes.Count -le 0) { continue }

                $boundLeft = [double](($boxes | Measure-Object -Property Left -Minimum).Minimum)
                $boundTop = [double](($boxes | Measure-Object -Property Top -Minimum).Minimum)
                $boundRight = [double](($boxes | Measure-Object -Property Right -Maximum).Maximum)
                $boundBottom = [double](($boxes | Measure-Object -Property Bottom -Maximum).Maximum)
                $boundWidth = $boundRight - $boundLeft
                $boundHeight = $boundBottom - $boundTop
                if ($boundWidth -le 1 -or $boundHeight -le 1) { continue }

                $isAnalysis = [bool]$spec.analysis
                $verticalShift = 0.0
                if (-not $isAnalysis -and [bool]$spec.inline) {
                    $verticalShift = __INLINE_SHIFT__
                } elseif (-not $isAnalysis -and (-not [bool]$spec.compact)) {
                    $verticalShift = __MULTILINE_SHIFT__
                }
                $targetLeft = [Math]::Max(0.0, $boundLeft - 2.0)
                $targetTop = [Math]::Max(0.0, $boundTop - $verticalShift - 2.0)
                $targetWidth = [Math]::Max(18.0, $boundWidth + 8.0)
                $targetHeight = [Math]::Max(20.0, $boundHeight + 8.0)
                if ([bool]$spec.compact -and [double]$spec.expectedWidth -gt 0) {
                    $targetWidth = [Math]::Max($targetWidth, [double]$spec.expectedWidth)
                    $targetHeight = [Math]::Min([Math]::Max($targetHeight, 20.0), 40.0)
                }
                try {
                    $slideWidth = [double]$pres.PageSetup.SlideWidth
                } catch {
                    $slideWidth = 960.0
                }
                $maxAllowedWidth = [Math]::Max(18.0, $slideWidth - $targetLeft - 8.0)
                $targetWidth = [Math]::Min($targetWidth, $maxAllowedWidth)

                $shape.TextFrame.MarginLeft = 0
                $shape.TextFrame.MarginRight = 0
                $shape.TextFrame.MarginTop = 0
                $shape.TextFrame.MarginBottom = 0
                try { $shape.TextFrame.WordWrap = -1 } catch {}
                $shape.TextFrame.TextRange.Text = [string]$spec.text
                try {
                    $paragraph = $shape.TextFrame.TextRange.ParagraphFormat
                    $paragraph.Alignment = 1
                    $paragraph.LeftIndent = 0
                    $paragraph.FirstLineIndent = 0
                    $paragraph.SpaceBefore = 0
                    $paragraph.SpaceAfter = 0
                } catch {}
                try {
                    $paragraph2 = $shape.TextFrame2.TextRange.ParagraphFormat
                    $paragraph2.Alignment = 1
                    $paragraph2.LeftIndent = 0
                    $paragraph2.FirstLineIndent = 0
                    $paragraph2.SpaceBefore = 0
                    $paragraph2.SpaceAfter = 0
                } catch {}
                $font = $shape.TextFrame.TextRange.Font
                $font.Name = "宋体"
                $font.NameFarEast = "宋体"
                $font.Size = 24
                $font.Bold = -1
                $font.Underline = 0
                $font.Color.RGB = 255
                if ($isAnalysis) {
                    foreach ($label in @("【关键能力】", "【解题思路】", "【解决思路】", "【解析】")) {
                        $searchFrom = 0
                        while ($true) {
                            $idx = [string]$shape.TextFrame.TextRange.Text
                            $idx = $idx.IndexOf($label, $searchFrom)
                            if ($idx -lt 0) { break }
                            try {
                                $labelRange = $shape.TextFrame.TextRange.Characters($idx + 1, $label.Length)
                                $labelFont = $labelRange.Font
                                $labelFont.Name = "黑体"
                                $labelFont.NameFarEast = "黑体"
                                $labelFont.Bold = -1
                            } catch {}
                            $searchFrom = $idx + $label.Length
                        }
                    }
                }
                $shape.Left = $targetLeft
                $shape.Top = $targetTop
                $shape.Width = $targetWidth
                $shape.Height = $targetHeight
                try {
                    $wrappedHeight = [double]$shape.TextFrame.TextRange.BoundHeight
                    if ($wrappedHeight -gt 1) {
                        $shape.Height = [Math]::Max([double]$shape.Height, $wrappedHeight + 8.0)
                    }
                } catch {}

                if ($shapeName -notmatch "answer-adjusted") {
                    $shape.Name = "$shapeName answer-adjusted"
                }
                $adjusted += 1
            } catch {
                continue
            }
        }
    }
    if ($adjusted -gt 0) { $pres.Save() }
    Write-Output "ANSWER_CROP_ADJUSTED=$adjusted"
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
    script = script.replace("__INLINE_SHIFT__", str(ANSWER_INLINE_VERTICAL_SHIFT_PT))
    script = script.replace("__MULTILINE_SHIFT__", str(ANSWER_MULTILINE_VERTICAL_SHIFT_PT))
    env = os.environ.copy()
    env["PPTX_TO_CROP"] = str(pptx)
    env["ANSWER_SPEC_JSON"] = str(spec_path)
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
        print(f"提示：PowerShell 答案文本框裁剪失败：{exc}", file=sys.stderr)
        return False
    try:
        spec_path.unlink()
    except OSError:
        pass
    if result.returncode != 0:
        stderr = result.stderr.strip() or result.stdout.strip()
        print(f"提示：PowerShell 答案文本框裁剪失败：{stderr}", file=sys.stderr)
        return False
    output = result.stdout.strip()
    if output:
        print(f"答案文本框裁剪（PowerShell）：{pptx.name}，{output}")
    return True


def postprocess_tif_positions_with_powerpoint(
    pptx: Path,
    *,
    adjust_answer_overlays: bool = True,
    multiline_answer_min_width_pt: float = 0.0,
) -> None:
    try:
        import win32com.client
    except Exception as exc:
        print(f"提示：未能加载 PowerPoint COM，TIF 位置使用 XML 估算值：{exc}", file=sys.stderr)
        if adjust_answer_overlays:
            if crop_answer_overlays_with_powershell(pptx):
                clear_answer_overlay_alpha(pptx)
        return

    answer_specs = collect_answer_overlay_specs(pptx) if adjust_answer_overlays else {}
    pp = None
    pres = None
    work_pptx = pptx
    temp_pptx: Path | None = None
    export_dir: Path | None = None
    export_index = 0
    if any(ord(ch) > 127 for ch in str(pptx)):
        temp_pptx = pptx.with_name(f"_ppt_com_work_{hashlib.sha1(str(pptx).encode('utf-8')).hexdigest()[:8]}.pptx")
        shutil.copy2(pptx, temp_pptx)
        work_pptx = temp_pptx
    try:
        pp = win32com.client.Dispatch("PowerPoint.Application")
        pp.Visible = True
        pres = pp.Presentations.Open(str(work_pptx), WithWindow=False)
        adjusted = 0
        answer_adjusted = 0
        animations_rebuilt = 0
        cover_styled = style_cover_subtitle_with_powerpoint(pres)
        for slide in pres.Slides:
            tif_pictures = []
            for shape in slide.Shapes:
                try:
                    if "fbd tif" in shape.Name.lower():
                        tif_pictures.append(shape)
                except Exception:
                    continue
            placeholders: list[tuple[float, float, float, float]] = []
            for shape in slide.Shapes:
                try:
                    if not (shape.HasTextFrame and shape.TextFrame.HasText):
                        continue
                    if "answer" in shape.Name.lower():
                        continue
                    text = shape.TextFrame.TextRange.Text
                except Exception:
                    continue
                if "\u3000" not in text:
                    continue
                for index, ch in enumerate(text, start=1):
                    if ord(ch) != 0x3000:
                        continue
                    try:
                        char_range = shape.TextFrame.TextRange.Characters(index, 1)
                        placeholders.append(
                            (
                                float(char_range.BoundLeft),
                                float(char_range.BoundTop),
                                float(char_range.BoundWidth),
                                float(char_range.BoundHeight),
                            )
                        )
                    except Exception:
                        continue

            long_answer_underlines = []
            long_answer_underline_index = 0
            if multiline_answer_min_width_pt > 0:
                for shape in slide.Shapes:
                    try:
                        lower_name = shape.Name.lower()
                        if "长答案横线" in shape.Name or "long answer underline" in lower_name:
                            long_answer_underlines.append(shape)
                    except Exception:
                        continue
                long_answer_underlines.sort(key=lambda item: (float(item.Top), float(item.Left)))

            for picture, (left, top, width, height) in zip(tif_pictures, placeholders):
                try:
                    picture.LockAspectRatio = -1
                except Exception:
                    pass
                try:
                    picture.Height = min(21.0, max(18.0, height * 0.60))
                    picture.Left = left + (width - picture.Width) / 2
                    picture.Top = top + (height - picture.Height) * 0.58
                    adjusted += 1
                except Exception:
                    continue

            for shape in slide.Shapes:
                try:
                    shape_name = shape.Name
                except Exception:
                    continue
                spec = answer_specs.get((int(slide.SlideIndex), shape_name))
                if spec is None:
                    continue
                visible_line_count_hint = len([line for line in spec.text_lines if line.strip()])
                text_hint = "".join(spec.text_lines).strip()
                is_analysis_hint = any(
                    label in text_hint
                    for label in ("【解析】", "【关键能力】", "【解题思路】", "【解决思路】")
                )
                if (
                    multiline_answer_min_width_pt > 0
                    and long_answer_underline_index < len(long_answer_underlines)
                    and not is_analysis_hint
                    and not spec.inline_context
                    and visible_line_count_hint >= 1
                ):
                    char_boxes_for_underlines = powerpoint_char_boxes_for_text(shape, spec.text_lines)
                    grouped_underline_lines: list[list[tuple[str, float, float, float, float]]] = []
                    for box in char_boxes_for_underlines:
                        for line_boxes in grouped_underline_lines:
                            if abs(box[2] - line_boxes[0][2]) <= 5:
                                line_boxes.append(box)
                                break
                        else:
                            grouped_underline_lines.append([box])
                    grouped_underline_lines.sort(key=lambda boxes_for_line: min(box[2] for box in boxes_for_line))
                    if grouped_underline_lines:
                        for line_boxes in grouped_underline_lines:
                            if long_answer_underline_index >= len(long_answer_underlines):
                                break
                            line_shape = long_answer_underlines[long_answer_underline_index]
                            long_answer_underline_index += 1
                            try:
                                text_bottom = max(box[4] for box in line_boxes)
                                line_shape.Top = max(0, text_bottom - 1.2)
                                adjusted += 1
                            except Exception:
                                continue
                if (
                    multiline_answer_min_width_pt > 0
                    and visible_line_count_hint > 1
                    and not is_analysis_hint
                    and not spec.compact
                ):
                    char_boxes = powerpoint_char_boxes_for_text(shape, spec.text_lines)
                    grouped_lines: list[list[tuple[str, float, float, float, float]]] = []
                    for box in char_boxes:
                        for line_boxes in grouped_lines:
                            if abs(box[2] - line_boxes[0][2]) <= 5:
                                line_boxes.append(box)
                                break
                        else:
                            grouped_lines.append([box])
                    grouped_lines.sort(key=lambda boxes_for_line: min(box[2] for box in boxes_for_line))
                    for line_boxes in grouped_lines:
                        if long_answer_underline_index >= len(long_answer_underlines):
                            break
                        line_shape = long_answer_underlines[long_answer_underline_index]
                        long_answer_underline_index += 1
                        try:
                            text_bottom = max(box[4] for box in line_boxes)
                            line_shape.Top = max(0, text_bottom - 1.2)
                            adjusted += 1
                        except Exception:
                            continue
                    continue
                visual_bounds = None
                base_char_boxes = []
                force_split_multiline_answer_fragments = False
                anchor_to_base_shape = should_anchor_answer_to_base_shape(shape_name) or (
                    not spec.inline_context and not spec.compact and not is_analysis_hint
                )
                if anchor_to_base_shape:
                    position_source_shape = find_answer_base_shape(slide, shape_name)
                    if position_source_shape is not None:
                        base_char_boxes = powerpoint_char_boxes_for_text(position_source_shape, spec.text_lines)
                        if spec.compact:
                            dominant_boxes = dominant_visual_line_char_boxes(base_char_boxes)
                            if dominant_boxes:
                                dominant_left = min(box[1] for box in dominant_boxes)
                                try:
                                    slide_width_pt = float(pres.PageSetup.SlideWidth)
                                except Exception:
                                    slide_width_pt = 960.0
                                if dominant_left + spec.expected_width_pt <= slide_width_pt - 6:
                                    base_char_boxes = dominant_boxes
                                else:
                                    force_split_multiline_answer_fragments = True
                if not spec.inline_context and not base_char_boxes:
                    try:
                        if export_dir is None:
                            export_dir = work_pptx.with_name(f"_shape_exports_{hashlib.sha1(str(work_pptx).encode('utf-8')).hexdigest()[:8]}")
                            export_dir.mkdir(parents=True, exist_ok=True)
                        export_index += 1
                        visual_bounds = exported_shape_visual_bounds(shape, export_dir / f"answer_{int(slide.SlideIndex)}_{export_index}.png")
                    except Exception:
                        visual_bounds = None
                if base_char_boxes:
                    char_boxes = base_char_boxes
                    boxes = [(box[1], box[2], box[3], box[4]) for box in char_boxes]
                elif spec.inline_context:
                    position_source_shape = find_answer_base_shape(slide, shape_name)
                    char_boxes = (
                        powerpoint_char_boxes_for_ranges(position_source_shape, spec.ranges)
                        if position_source_shape is not None
                        else []
                    )
                    if not char_boxes and position_source_shape is not None:
                        char_boxes = powerpoint_char_boxes_for_text(position_source_shape, spec.text_lines)
                    if char_boxes:
                        boxes = [(box[1], box[2], box[3], box[4]) for box in char_boxes]
                    else:
                        try:
                            if export_dir is None:
                                export_dir = work_pptx.with_name(f"_shape_exports_{hashlib.sha1(str(work_pptx).encode('utf-8')).hexdigest()[:8]}")
                                export_dir.mkdir(parents=True, exist_ok=True)
                            export_index += 1
                            visual_bounds = exported_shape_visual_bounds(shape, export_dir / f"inline_answer_{int(slide.SlideIndex)}_{export_index}.png")
                        except Exception:
                            visual_bounds = None
                        if visual_bounds is not None:
                            char_boxes = []
                            boxes = [visual_bounds]
                        else:
                            char_boxes = powerpoint_char_boxes_for_ranges(shape, spec.ranges)
                            if not char_boxes:
                                char_boxes = powerpoint_char_boxes_for_text(shape, spec.text_lines)
                            boxes = [(box[1], box[2], box[3], box[4]) for box in char_boxes]
                elif visual_bounds is not None:
                    char_boxes = []
                    boxes = [visual_bounds]
                else:
                    char_boxes = powerpoint_char_boxes_for_text(shape, spec.text_lines)
                    if not char_boxes:
                        char_boxes = powerpoint_char_boxes_for_ranges(shape, spec.ranges)
                    boxes = [(box[1], box[2], box[3], box[4]) for box in char_boxes]
                if not boxes:
                    continue
                line_tops: list[float] = []
                for box in boxes:
                    if all(abs(box[1] - top_value) > 5 for top_value in line_tops):
                        line_tops.append(box[1])
                visible_line_count = len([line for line in spec.text_lines if line.strip()])
                is_analysis = any(
                    label in "".join(spec.text_lines)
                    for label in ("【解析】", "【关键能力】", "【解题思路】", "【解决思路】")
                )
                is_single_visual_line = len(line_tops) <= 1
                left = min(box[0] for box in boxes)
                top = min(box[1] for box in boxes)
                right = max(box[2] for box in boxes)
                bottom = max(box[3] for box in boxes)
                text = "\r".join(line for line in spec.text_lines if line)
                if not text:
                    continue
                split_multiline_answer_fragments = force_split_multiline_answer_fragments or (
                    spec.inline_context and len(line_tops) > 1 and not is_analysis
                )
                if split_multiline_answer_fragments and len(line_tops) > 1 and not is_analysis:
                    fragments: list[tuple[str, float, float, float, float]] = []
                    for line_top in sorted(line_tops):
                        line_chars = [box for box in char_boxes if abs(box[2] - line_top) <= 5]
                        line_chars.sort(key=lambda item: item[1])
                        fragment_text = "".join(box[0] for box in line_chars).strip(" \t\r\n\u3000_（）()")
                        if not fragment_text:
                            continue
                        fragment_left = min(box[1] for box in line_chars)
                        fragment_top = min(box[2] for box in line_chars)
                        fragment_right = max(box[3] for box in line_chars)
                        fragment_bottom = max(box[4] for box in line_chars)
                        fragments.append((fragment_text, fragment_left, fragment_top, fragment_right, fragment_bottom))
                    if fragments:
                        for fragment_index, (fragment_text, fragment_left, fragment_top, fragment_right, fragment_bottom) in enumerate(fragments):
                            try:
                                fragment_anchor_top = fragment_top - (3.0 if spec.inline_context else 0.0)
                                target_shape = shape if fragment_index == 0 else shape.Duplicate().Item(1)
                                target_shape.TextFrame.MarginLeft = 0
                                target_shape.TextFrame.MarginRight = 0
                                target_shape.TextFrame.MarginTop = 0
                                target_shape.TextFrame.MarginBottom = 0
                                try:
                                    target_shape.TextFrame.WordWrap = -1
                                except Exception:
                                    pass
                                target_shape.TextFrame.TextRange.Text = fragment_text
                                try:
                                    paragraph = target_shape.TextFrame.TextRange.ParagraphFormat
                                    paragraph.Alignment = 1
                                    paragraph.LeftIndent = 0
                                    paragraph.FirstLineIndent = 0
                                    paragraph.SpaceBefore = 0
                                    paragraph.SpaceAfter = 0
                                except Exception:
                                    pass
                                try:
                                    paragraph2 = target_shape.TextFrame2.TextRange.ParagraphFormat
                                    paragraph2.Alignment = 1
                                    paragraph2.LeftIndent = 0
                                    paragraph2.FirstLineIndent = 0
                                    paragraph2.SpaceBefore = 0
                                    paragraph2.SpaceAfter = 0
                                except Exception:
                                    pass
                                font = target_shape.TextFrame.TextRange.Font
                                font.Name = "宋体"
                                font.NameFarEast = "宋体"
                                font.Size = 24
                                font.Bold = True
                                font.Underline = False
                                font.Color.RGB = 255
                                target_shape.Left = max(0, fragment_left - 2)
                                target_shape.Top = max(0, fragment_anchor_top - 2)
                                target_shape.Width = max(18, fragment_right - fragment_left + 6)
                                target_shape.Height = max(20, fragment_bottom - fragment_top + 6)
                                align_shape_text_to_bounds(target_shape, fragment_left, max(0, fragment_anchor_top))
                                target_shape.Name = f"{shape_name} fragment {fragment_index + 1} answer-adjusted"
                                answer_adjusted += 1
                            except Exception:
                                continue
                        continue
                answer_vertical_shift = 0.0
                if not is_analysis:
                    if spec.inline_context:
                        answer_vertical_shift = ANSWER_INLINE_VERTICAL_SHIFT_PT
                    elif not spec.compact and not base_char_boxes:
                        # Long-answer blanks are rendered as editable text
                        # underlines in the base shape.  Their red answer
                        # overlays need a small lift so the underline sits
                        # beneath the glyphs instead of cutting through them.
                        answer_vertical_shift = ANSWER_MULTILINE_VERTICAL_SHIFT_PT
                anchor_top = max(0, top - answer_vertical_shift)
                target_left = max(0, left - 2)
                target_top = max(0, anchor_top - 2)
                target_width = max(18, right - left + 8)
                target_height = max(24, bottom - top + 8)
                if not is_analysis:
                    try:
                        slide_width_pt = float(pres.PageSetup.SlideWidth)
                    except Exception:
                        slide_width_pt = 960.0
                    max_allowed_width = max(18.0, slide_width_pt - target_left - 8.0)
                    target_width = min(max_allowed_width, target_width)
                if (spec.compact or is_single_visual_line) and target_width > spec.expected_width_pt * 1.35:
                    target_width = spec.expected_width_pt
                    first_left = boxes[0][0]
                    slide_width_pt = float(pres.PageSetup.SlideWidth)
                    if first_left + target_width < slide_width_pt - 6:
                        target_left = max(0, first_left - 2)
                    target_height = max(20, min(target_height, 40))
                if spec.compact and target_width < spec.expected_width_pt:
                    target_width = spec.expected_width_pt
                    target_height = max(20, min(target_height, 40))
                try:
                    shape.TextFrame.MarginLeft = 0
                    shape.TextFrame.MarginRight = 0
                    shape.TextFrame.MarginTop = 0
                    shape.TextFrame.MarginBottom = 0
                    try:
                        shape.TextFrame.WordWrap = -1
                    except Exception:
                        pass
                    shape.TextFrame.TextRange.Text = text
                    try:
                        paragraph = shape.TextFrame.TextRange.ParagraphFormat
                        paragraph.Alignment = 1
                        paragraph.LeftIndent = 0
                        paragraph.FirstLineIndent = 0
                        paragraph.SpaceBefore = 0
                        paragraph.SpaceAfter = 0
                    except Exception:
                        pass
                    try:
                        paragraph2 = shape.TextFrame2.TextRange.ParagraphFormat
                        paragraph2.Alignment = 1
                        paragraph2.LeftIndent = 0
                        paragraph2.FirstLineIndent = 0
                        paragraph2.SpaceBefore = 0
                        paragraph2.SpaceAfter = 0
                    except Exception:
                        pass
                    font = shape.TextFrame.TextRange.Font
                    font.Name = "宋体"
                    font.NameFarEast = "宋体"
                    font.Size = 24
                    font.Bold = True
                    font.Underline = False
                    font.Color.RGB = 255
                    if is_analysis:
                        full_text = shape.TextFrame.TextRange.Text
                        for label in ("【关键能力】", "【解题思路】", "【解决思路】", "【解析】"):
                            search_from = 0
                            while True:
                                idx = full_text.find(label, search_from)
                                if idx < 0:
                                    break
                                try:
                                    label_range = shape.TextFrame.TextRange.Characters(idx + 1, len(label))
                                    label_font = label_range.Font
                                    label_font.Name = "黑体"
                                    label_font.NameFarEast = "黑体"
                                    label_font.Bold = True
                                except Exception:
                                    pass
                                search_from = idx + len(label)
                    shape.Left = target_left
                    shape.Top = target_top
                    shape.Width = target_width
                    shape.Height = target_height
                    try:
                        wrapped_height = float(shape.TextFrame.TextRange.BoundHeight)
                        if wrapped_height > 1:
                            shape.Height = max(float(shape.Height), wrapped_height + 8)
                    except Exception:
                        pass
                    align_shape_text_to_bounds(shape, left, anchor_top)
                    if "answer-adjusted" not in shape.Name.lower():
                        shape.Name = f"{shape.Name} answer-adjusted"
                    answer_adjusted += 1
                except Exception:
                    continue
            animations_rebuilt += add_powerpoint_standard_animations(slide)
        if adjusted:
            print(f"TIF 精准校位：{pptx.name}，已调整 {adjusted} 处")
        if answer_adjusted:
            print(f"答案文本框裁剪：{pptx.name}，已调整 {answer_adjusted} 处")
        if animations_rebuilt:
            print(f"答案动画重建：{pptx.name}，已重建 {animations_rebuilt} 个对象")
        if cover_styled:
            print(f"首页副标题字体：{pptx.name}，已设置 {cover_styled} 处")
        if adjusted or answer_adjusted or animations_rebuilt or cover_styled:
            pres.Save()
    except Exception as exc:
        print(f"提示：PowerPoint TIF 精准校位失败，保留 XML 估算位置：{exc}", file=sys.stderr)
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if pp is not None:
            try:
                pp.Quit()
            except Exception:
                pass
        if export_dir is not None and export_dir.exists():
            try:
                shutil.rmtree(export_dir)
            except Exception:
                pass
        if temp_pptx is not None and temp_pptx.exists():
            try:
                shutil.copy2(temp_pptx, pptx)
            finally:
                try:
                    temp_pptx.unlink()
                except Exception:
                    pass
    if adjust_answer_overlays:
        clear_answer_overlay_alpha(pptx)


def build_deck(
    template: Path,
    output: Path,
    pieces: list[Piece],
    answer_map: dict[tuple[str, int], Piece],
    volume: str,
    start: int,
    end: int,
    *,
    tif_search_dirs: list[Path],
    multiline_answer_min_width_pt: float = 0.0,
    content_src_override: int | None = None,
) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    slide_specs: list[tuple[int, object]] = []

    # Source slide contract for 教材文言梳理模板:
    # 1 cover, 2 instruction, 3 directory, 4 title, 5 blank content, last 3 preserve.
    with zipfile.ZipFile(template, "r") as zin:
        slide_count_template = len([n for n in zin.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)])
        title_src, content_src, final_sources, content_box = template_contract(slide_count_template)
        if content_src_override is not None:
            content_src = content_src_override

        generated_slides: list[tuple[int, bytes, bytes]] = []
        extra_media_assets: dict[str, bytes] = {}
        piece_chunks: list[tuple[Piece, list[list[StyledLine]]]] = []
        title_slide_numbers: dict[int, int] = {}
        next_generated_slide = 4
        for piece in pieces:
            chunks = make_piece_slide_chunks(piece, answer_map.get(piece_key(piece)))
            piece_chunks.append((piece, chunks))
            title_slide_numbers[piece.num] = next_generated_slide
            next_generated_slide += 1 + len(chunks)

        def add_from(src: int, edit_func=lambda root, rels_root: None) -> None:
            xml, rels = make_slide_parts(zin, src, edit_func)
            generated_slides.append((src, xml, rels))

        add_from(1, lambda root, rels_root: add_cover_content(root, volume, start, end))
        add_from(2)
        add_from(3, lambda root, rels_root: add_directory_content(root, rels_root, pieces, title_slide_numbers))

        for piece, chunks in piece_chunks:
            add_from(title_src, lambda root, rels_root, piece=piece: add_piece_title(root, piece))
            for chunk in chunks:
                add_from(
                    content_src,
                    lambda root, rels_root, piece=piece, chunk=chunk: add_content(
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
        content_types = update_content_types(parse_xml(zin.read("[Content_Types].xml")), slide_total)
        presentation_xml = update_presentation_xml(parse_xml(zin.read("ppt/presentation.xml")), slide_total)
        presentation_rels = update_presentation_rels(parse_xml(zin.read("ppt/_rels/presentation.xml.rels")), slide_total)

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
            for module in MODULE_IMAGE_FILES:
                image_info = module_image_package(module)
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
            for i, (_src, xml, rels) in enumerate(generated_slides, 1):
                zout.writestr(f"ppt/slides/slide{i}.xml", xml)
                zout.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", rels)
    postprocess_tif_positions_with_powerpoint(output, multiline_answer_min_width_pt=multiline_answer_min_width_pt)


def find_md_files(run_dir: Path) -> tuple[Path, Path]:
    body = sorted(run_dir.glob("正文_*.md"))
    answer = sorted(run_dir.glob("答案_*.md"))
    if not body or not answer:
        raise FileNotFoundError(f"未在目录中找到正文_*.md 和 答案_*.md：{run_dir}")
    return body[0], answer[0]


def infer_tif_search_dirs(run_dir: Path) -> list[Path]:
    dirs: list[Path] = []
    manifest_path = run_dir / "conversion_manifest.json"
    if manifest_path.exists():
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            for item in manifest.get("outputs", []):
                source = item.get("source")
                if source:
                    dirs.append(Path(source).parent)
        except Exception:
            pass
    dirs.append(run_dir)
    result: list[Path] = []
    seen: set[str] = set()
    for directory in dirs:
        key = str(directory).lower()
        if key not in seen and directory.exists():
            result.append(directory)
            seen.add(key)
    return result


def output_file_name_for_group(group_index: int, file_name: str) -> str:
    number = str(group_index)
    if group_index in {11, 14} and file_name.startswith(number) and not file_name.startswith(f"{number} "):
        return f"{number} {file_name[len(number):]}"
    return file_name


def generate_wenyan_pptx(
    *,
    md_run: Path,
    template: Path,
    output_dir: Path,
    limit_groups: int | None = None,
    group_ranges: list[dict[str, object]] | None = None,
    progress: Callable[[str], None] | None = None,
) -> list[Path]:
    def emit(message: str) -> None:
        if progress:
            progress(message)
        else:
            print(message)

    body_md, answer_md = find_md_files(md_run)
    emit(f"正文 MD：{body_md}")
    emit(f"答案 MD：{answer_md}")
    emit(f"模板 PPTX：{template}")

    body_pieces = parse_pieces(body_md, is_answer=False)
    answer_pieces = parse_pieces(answer_md, is_answer=True)
    answer_map = {piece_key(piece): piece for piece in answer_pieces}
    tif_search_dirs = infer_tif_search_dirs(md_run)

    emit(f"正文篇目数：{len(body_pieces)}")
    emit(f"答案篇目数：{len(answer_pieces)}")

    selected_groups = GROUPS if limit_groups is None else GROUPS[:limit_groups]
    if group_ranges is not None:
        selected_groups = []
        for index, group in enumerate(group_ranges, 1):
            volume = str(group["volume"])
            start = int(group["start"])
            end = int(group["end"])
            keys = {tuple(key) for key in group["unit_keys"]}  # type: ignore[arg-type]
            members = [piece for piece in body_pieces if (piece.volume, piece.num) in keys]
            selected_groups.append((f"{index}{volume} 第{start}-{end}篇.pptx", volume, start, end, members))
    else:
        selected_groups = [(*group, None) for group in selected_groups]
    output_dir.mkdir(parents=True, exist_ok=True)
    generated: list[Path] = []
    for group_index, (file_name, volume, start, end, selected_override) in enumerate(selected_groups, 1):
        final_file_name = output_file_name_for_group(group_index, file_name)
        group_pieces = selected_override or [p for p in body_pieces if p.volume == volume and start <= p.num <= end]
        if not group_pieces:
            emit(f"跳过，无篇目：{final_file_name}")
            continue
        output = output_dir / final_file_name
        emit(f"[{group_index}/{len(selected_groups)}] 生成：{output.name}")
        build_deck(template, output, group_pieces, answer_map, volume, start, end, tif_search_dirs=tif_search_dirs)
        stale = output_dir / file_name
        if stale != output and stale.exists():
            stale.unlink()
        generated.append(output)
        emit(f"已生成：{output}，篇目：{', '.join(f'{p.num}.{p.title}' for p in group_pieces)}")

    manifest = {
        "md_run": str(md_run),
        "template": str(template),
        "generated": [str(path) for path in generated],
    }
    (output_dir / "test_generation_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return generated


def scan_wenyan_group_units(md_run: Path) -> list[dict[str, object]]:
    """Return the detected volume/piece sequence for dynamic GUI grouping."""

    body_md, _answer_md = find_md_files(md_run)
    return [
        {"volume": piece.volume, "num": piece.num, "title": piece.title}
        for piece in parse_pieces(body_md, is_answer=False)
    ]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="测试：根据正文/答案 md 和模板生成前 1-2 个 PPTX。")
    parser.add_argument("--md-run", default=str(DEFAULT_MD_RUN), help="GUI 生成的 md run 目录。")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="PPTX 模板路径。")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="测试 PPTX 输出目录。")
    parser.add_argument("--limit-groups", type=int, default=2, help="生成前几个分组，默认 2；传 0 表示全部。")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    run_dir = Path(args.md_run)
    template = Path(args.template)
    output_dir = Path(args.output_dir)
    generate_wenyan_pptx(
        md_run=run_dir,
        template=template,
        output_dir=output_dir,
        limit_groups=None if args.limit_groups == 0 else args.limit_groups,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

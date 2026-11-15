#!/usr/bin/env python3
"""Test generator for 语文《文言文阅读加分练》.

This adapter keeps the book-specific parsing thin and reuses the mature
三步进阶/文言梳理 rendering stack for slide construction, answer overlays,
animations, long underline answers, and TIF placement.
"""

from __future__ import annotations

import argparse
import copy
import json
import math
import re
import shutil
import sys
import tempfile
import zipfile
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Callable
from xml.etree import ElementTree as ET


PROJECT_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
TESTS_DIR = Path(__file__).resolve().parent
TOOLS_DIR = PROJECT_DIR / "tools"
if str(TESTS_DIR) not in sys.path:
    sys.path.insert(0, str(TESTS_DIR))
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

import test_md_to_ppt as wenyan  # noqa: E402
import test_sanbu_md_to_ppt as sanbu  # noqa: E402
from fbd_to_md import convert_fbd_to_markdown, read_fbd, tif_markdown_tag  # noqa: E402
from generation_state import isolated_generation  # noqa: E402


DEFAULT_SOURCE_DIR = PROJECT_DIR / "文言文阅读加分练_源文件"
DEFAULT_BODY_FBD = DEFAULT_SOURCE_DIR / "文言文阅读加分练正文.fbd"
DEFAULT_ANSWER_FBD = DEFAULT_SOURCE_DIR / "文言文阅读加分练答案.fbd"
DEFAULT_TEMPLATE = PROJECT_DIR / "文言文阅读加分练模板.pptx"
DEFAULT_MD_ROOT = PROJECT_DIR / "out_md" / "文言文阅读加分练"
DEFAULT_OUTPUT_DIR = PROJECT_DIR / "output" / "test_wenyan_yuedu_jiafen_ppt"

YUEDU_GROUPS = [
    ("1真题速递  篇目一至二.pptx", "真题速递", 1, 2),
    ("2专题一    论说类(篇目一至五).pptx", "专题一论说类", 1, 5),
    ("3专题一    论说类(篇目六至十).pptx", "专题一论说类", 6, 10),
    ("4专题二    人物传记类(篇目一至五).pptx", "专题二人物传记类", 1, 5),
    ("5专题三    记事类(篇目一至三).pptx", "专题三记事类", 1, 3),
    ("6专题三    记事类(篇目四至六).pptx", "专题三记事类", 4, 6),
    ("7专题四    游记散文类(篇目一至四).pptx", "专题四游记散文类", 1, 4),
    ("8专题五    其他类(篇目一至三).pptx", "专题五其他类", 1, 3),
    ("9专题五    其他类(篇目四至五).pptx", "专题五其他类", 4, 5),
]

VOLUME_MARKER_RE = re.compile(r"^(真题速递|专题[一二三四五六七八九十]+.*)$")
PIECE_TITLE_RE = re.compile(r"^([一二三四五六七八九十]+)、(.+)$")
MAIN_SECTION_TITLES = {"深研文本", "对标真题", "教考衔接", "联读拓展", "文化常识"}
REFERENCE_TIF_NAMES = {"参考译文.tif", "知识拓展.tif"}
BG_START = "__YUEDU_BG_START__"
BG_END = "__YUEDU_BG_END__"
MATERIAL_LABEL_RE = re.compile(r"^材料[一二三四五六七八九十]+[:：]$")
ANALYSIS_LABEL_RE = re.compile(r"(【(?:关键能力|解题思路|解决思路)】)")
YUEDU_LONG_ANSWER_TEXT_UNITS = 30.0
YUEDU_LONG_ANSWER_UNDERLINE_UNITS = 34.0
YUEDU_LONG_ANSWER_TRAILING_UNITS = 3.0
CHINESE_NUMBERS = {
    "一": 1,
    "二": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
    "十": 10,
}
ARABIC_NUMBERS = {value: key for key, value in CHINESE_NUMBERS.items() if value <= 10}
ARABIC_NUMBERS.update({11: "十一", 12: "十二", 13: "十三", 14: "十四", 15: "十五"})


def plain(text: str) -> str:
    return sanbu.plain(text)


def clean_lines(path: Path) -> list[str]:
    return sanbu.clean_sanbu_lines(path)


def chinese_number(value: str) -> int:
    value = value.strip()
    if value in CHINESE_NUMBERS:
        return CHINESE_NUMBERS[value]
    if value.startswith("十"):
        return 10 + CHINESE_NUMBERS.get(value[1:], 0)
    if value.endswith("十"):
        return CHINESE_NUMBERS.get(value[:-1], 1) * 10
    if "十" in value:
        left, right = value.split("十", 1)
        return CHINESE_NUMBERS.get(left, 1) * 10 + CHINESE_NUMBERS.get(right, 0)
    return 0


def normalize_title(text: str) -> str:
    return re.sub(r"\s+", "", plain(text)).strip()


def normalize_volume(text: str) -> str:
    return re.sub(r"[\s_＿]+", "", plain(text)).strip()


def is_volume_marker(text: str) -> bool:
    return bool(VOLUME_MARKER_RE.fullmatch(normalize_volume(text)))


def title_key(text: str) -> str:
    value = normalize_title(text)
    value = re.sub(r"[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]+$", "", value)
    return value


def strip_piece_prefix(title: str) -> str:
    return PIECE_TITLE_RE.sub(r"\2", plain(title)).strip()


def collect_tif_keep_map(source_dir: Path) -> dict[str, str]:
    return {path.name.lower(): tif_markdown_tag(path.name) for path in source_dir.glob("*.tif")}


def preserve_yuedu_background_markers(raw: str) -> str:
    """Preserve FBD =D highlighted terms for this book's 深研文本 glossary task."""

    return re.sub(r"〖=D\(〗(.*?)〖=〗", rf"{BG_START}\1{BG_END}", raw, flags=re.S)


def restore_yuedu_background_spans(markdown: str) -> str:
    return markdown.replace(BG_START, '<span class="yuedu-bg-word">').replace(BG_END, "</span>")


def convert_wenyan_yuedu_fbd_pair(
    *,
    body_fbd: Path = DEFAULT_BODY_FBD,
    answer_fbd: Path = DEFAULT_ANSWER_FBD,
    output_dir: Path | None = None,
    progress: Callable[[str], None] | None = None,
) -> Path:
    """Convert the two source FBD files into the MD run shape used by tests."""

    def emit(message: str) -> None:
        if progress:
            progress(message)
        else:
            print(message)

    run_dir = output_dir or DEFAULT_MD_ROOT / datetime.now().strftime("run_%Y%m%d_%H%M%S")
    run_dir.mkdir(parents=True, exist_ok=True)
    keep_map = collect_tif_keep_map(body_fbd.parent)
    emit(f"文言文阅读加分练：自动保留 TIF 资源 {len(keep_map)} 个。")

    outputs: list[dict[str, object]] = []
    for role, safe_name, fbd_path in [("正文", "body", body_fbd), ("答案", "answer", answer_fbd)]:
        emit(f"转换{role} FBD：{fbd_path}")
        raw = preserve_yuedu_background_markers(read_fbd(fbd_path, "gb18030"))
        markdown, stats = convert_fbd_to_markdown(raw, tif_replacements=keep_map, auto_start=False)
        markdown = restore_yuedu_background_spans(markdown)
        md_path = run_dir / f"{role}_{safe_name}.md"
        md_path.write_text(markdown, encoding="utf-8", newline="\n")
        outputs.append({"role": role, "source": str(fbd_path), "markdown": str(md_path), "stats": asdict(stats)})
        emit(f"{role} MD：{md_path.name}，TIF 保留 {stats.tif_kept}，未知 {stats.tif_unknown}")

    (run_dir / "conversion_manifest.json").write_text(
        json.dumps({"created_at": datetime.now().isoformat(timespec="seconds"), "outputs": outputs}, ensure_ascii=False, indent=2),
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


def parse_answer_piece_starts(lines: list[str]) -> list[tuple[int, str, int, str]]:
    starts: list[tuple[int, str, int, str]] = []
    current_volume = ""
    for index, line in enumerate(lines):
        value = plain(line)
        if is_volume_marker(value):
            current_volume = normalize_volume(value)
            continue
        match = PIECE_TITLE_RE.match(value)
        if current_volume and match:
            num = chinese_number(match.group(1))
            if num:
                starts.append((index, current_volume, num, value))
    return starts


def split_answer_sections(lines: list[str]) -> tuple[list[str], list[str]]:
    """Separate reference translation without swallowing later answer sections."""

    main: list[str] = []
    reference: list[str] = []
    in_reference = False
    for line in lines:
        value = plain(line)
        lower = line.lower()
        if any(name.lower() in lower for name in REFERENCE_TIF_NAMES):
            in_reference = True
            reference.append(line)
            continue
        if not value:
            continue
        if value == "答案解析" or is_volume_marker(value):
            continue
        if in_reference and value in {"教考衔接", "联读拓展", "文化常识"}:
            in_reference = False
            main.append(line)
            continue
        if in_reference:
            reference.append(line)
        else:
            main.append(line)
    return main, reference


def parse_answer_pieces(answer_md: Path) -> list[wenyan.Piece]:
    lines = clean_lines(answer_md)
    starts = parse_answer_piece_starts(lines)
    pieces: list[wenyan.Piece] = []
    for idx, (start, volume, num, title) in enumerate(starts):
        end = starts[idx + 1][0] if idx + 1 < len(starts) else len(lines)
        block = lines[start + 1 : end]
        main, reference = split_answer_sections(block)
        clean_title = strip_piece_prefix(title)
        piece_lines = [clean_title, sanbu.MAIN_MARKER, *main]
        if reference:
            piece_lines.extend([sanbu.REFERENCE_MARKER, *reference])
        pieces.append(wenyan.Piece(volume=volume, num=num, title=clean_title, author="", lines=piece_lines))
    return pieces


def body_volume_marker_positions(lines: list[str]) -> list[int]:
    return [index for index, line in enumerate(lines) if is_volume_marker(line)]


def body_title_candidates(lines: list[str], answer_title: str) -> list[tuple[int, str]]:
    expected = title_key(answer_title)
    candidates: list[tuple[int, str]] = []
    for index, line in enumerate(lines):
        value = plain(line)
        match = PIECE_TITLE_RE.match(value)
        if not match:
            continue
        candidate_title = strip_piece_prefix(value)
        candidate_key = title_key(candidate_title)
        if candidate_key == expected or candidate_key.startswith(expected):
            candidates.append((index, value))
    return candidates


def parse_body_pieces(body_md: Path, answer_pieces: list[wenyan.Piece]) -> list[wenyan.Piece]:
    lines = clean_lines(body_md)
    starts: list[tuple[int, wenyan.Piece, str]] = []
    used_positions: set[int] = set()
    search_from = 0
    for answer_piece in answer_pieces:
        candidates = body_title_candidates(lines, answer_piece.title)
        start = next((pos for pos, _title in candidates if pos >= search_from and pos not in used_positions), None)
        if start is None:
            start = next((pos for pos, _title in candidates if pos not in used_positions), None)
        if start is None:
            continue
        used_positions.add(start)
        search_from = start + 1
        starts.append((start, answer_piece, plain(lines[start])))

    volume_boundaries = body_volume_marker_positions(lines)
    pieces: list[wenyan.Piece] = []
    for idx, (start, answer_piece, raw_title) in enumerate(starts):
        next_piece_start = starts[idx + 1][0] if idx + 1 < len(starts) else len(lines)
        next_volume = next((pos for pos in volume_boundaries if pos > start), len(lines))
        end = min(next_piece_start, next_volume)
        content_lines = [line for line in lines[start + 1 : end] if line.strip()]
        pieces.append(
            wenyan.Piece(
                volume=answer_piece.volume,
                num=answer_piece.num,
                title=strip_piece_prefix(raw_title) or answer_piece.title,
                author="",
                lines=[answer_piece.title, sanbu.MAIN_MARKER, *content_lines],
            )
        )
    return pieces


def scan_yuedu_group_units(md_run: Path) -> list[dict[str, object]]:
    """Expose detected topics and passages for user-defined PPT ranges."""

    body_md, answer_md = find_md_files(md_run)
    answer_pieces = parse_answer_pieces(answer_md)
    body_pieces = parse_body_pieces(body_md, answer_pieces)
    return [
        {"volume": piece.volume, "num": piece.num, "title": piece.title}
        for piece in body_pieces
    ]


def custom_yuedu_filename(index: int, volume: str, start: int, end: int) -> str:
    piece_range = cover_piece_range(start, end)
    if volume == "真题速递":
        return f"{index}真题速递  {piece_range}.pptx"
    match = re.match(r"^(专题[一二三四五六七八九十]+)(.*)$", volume)
    if match:
        return f"{index}{match.group(1)}    {match.group(2)}({piece_range}).pptx"
    return f"{index}{volume}({piece_range}).pptx"


def split_yuedu_main_sections(lines: list[str], *, answer: bool = False) -> dict[str, list[str]]:
    sections: dict[str, list[str]] = {}
    current = "深研文本"
    buffer: list[str] = []

    def flush() -> None:
        nonlocal buffer
        if buffer:
            sections.setdefault(current, []).extend(buffer)
        buffer = []

    for line in lines:
        value = plain(line)
        if value == "深研文本":
            flush()
            current = "深研文本"
            if not answer:
                buffer.append(line)
            continue
        if value in MAIN_SECTION_TITLES and value != "深研文本":
            flush()
            current = value
            if not answer:
                buffer.append(line)
            continue
        buffer.append(line)
    flush()
    return sections


def chunk_units(chunk: list[wenyan.StyledLine]) -> float:
    return sum(wenyan.text_units(plain(line.text)) for line in chunk)


def is_sparse_intro_chunk(chunk: list[wenyan.StyledLine]) -> bool:
    if not chunk:
        return False
    animated_roles = {"answer", "analysis", "analysis_block", "segmentation_analysis", "long_answer"}
    if any(line.role in animated_roles for line in chunk):
        return False
    visible = [plain(line.text) for line in chunk if plain(line.text)]
    return 0 < len(visible) <= 3 and chunk_units(chunk) <= 18


def merge_sparse_intro_chunks(chunks: list[list[wenyan.StyledLine]]) -> list[list[wenyan.StyledLine]]:
    result: list[list[wenyan.StyledLine]] = []
    index = 0
    while index < len(chunks):
        current = chunks[index]
        if index + 1 < len(chunks) and is_sparse_intro_chunk(current):
            result.append(current + chunks[index + 1])
            index += 2
            continue
        result.append(current)
        index += 1
    return result


def split_text_by_units(text: str, max_units: float = 42.0) -> list[str]:
    pieces: list[str] = []
    current = ""
    units = 0.0
    last_break = -1
    break_chars = "。；！？;，,"
    for ch in text:
        weight = 1.0 if ord(ch) > 127 else 0.55
        if current and units + weight > max_units:
            if last_break >= 0 and wenyan.text_units(current[: last_break + 1]) >= max_units * 0.45:
                pieces.append(current[: last_break + 1].strip())
                current = current[last_break + 1 :].lstrip()
            else:
                pieces.append(current.strip())
                current = ""
            units = wenyan.text_units(current)
            last_break = -1
        current += ch
        units += weight
        if ch in break_chars:
            last_break = len(current) - 1
    if current.strip():
        pieces.append(current.strip())
    return pieces or [text]


def split_long_analysis_lines(lines: list[wenyan.StyledLine]) -> list[wenyan.StyledLine]:
    result: list[wenyan.StyledLine] = []
    split_roles = {"analysis", "analysis_block", "segmentation_analysis"}
    for line in lines:
        if line.role not in split_roles or wenyan.text_units(plain(line.text)) <= 56:
            result.append(line)
            continue
        for piece in split_text_by_units(line.text):
            result.append(wenyan.StyledLine(piece, line.role, line.meta))
    return result


def style_analysis_label_runs(text: str) -> str:
    return ANALYSIS_LABEL_RE.sub(r'<span class="font-heiti bold-text">\1</span>', text)


def is_material_label_text(text: str) -> bool:
    return bool(MATERIAL_LABEL_RE.fullmatch(plain(text).strip()))


def is_choice_question_group(group_lines: list[str]) -> bool:
    letters = {
        plain(line).strip()[:1]
        for line in group_lines
        if re.match(r"^[A-D][\.．、]", plain(line).strip())
    }
    return {"A", "B", "C", "D"}.issubset(letters)


def is_choice_analysis_chunk(lines: list[wenyan.StyledLine]) -> bool:
    analysis_index = next(
        (
            index
            for index, line in enumerate(lines)
            if (line.role or "") in {"analysis", "analysis_block", "segmentation_analysis", "yuedu_analysis"}
        ),
        None,
    )
    if analysis_index is None:
        return False
    body_lines = [
        line.text
        for line in lines[:analysis_index]
        if (line.role or "body") in {"body", "reading", "extended_reading"}
    ]
    return is_choice_question_group(body_lines)


def highlight_sup_target_words(line: str) -> str:
    # The glossary task marks the explained word immediately before [n].
    # Some FBD runs lose the original fill marker, so rebuild it from that stable cue.
    return re.sub(
        r"(?<!>)([\u4e00-\u9fff])(<sup>\[\d+\]</sup>)",
        r'<span class="yuedu-bg-word">\1</span>\2',
        line,
    )


def dot_terms_to_background(line: str) -> str:
    """Some short highlighted terms arrive as emphasis-dot spans in reading text."""

    return re.sub(
        r'<span\s+class="dot"[^>]*>(.*?)</span>',
        r'<span class="yuedu-bg-word">\1</span>',
        line,
        flags=re.S,
    )


def span_terms(line: str, class_name: str) -> list[str]:
    pattern = rf'<span\s+class="{re.escape(class_name)}"[^>]*>(.*?)</span>'
    return [plain(match.group(1)).strip() for match in re.finditer(pattern, line, flags=re.S) if plain(match.group(1)).strip()]


def glossary_prompt_index(lines: list[str]) -> int | None:
    for index, line in enumerate(lines):
        if is_deep_text_glossary_title(line):
            return index
    return None


def deep_glossary_question_terms(lines: list[str]) -> list[str]:
    start = glossary_prompt_index(lines)
    if start is None:
        return []
    terms: list[str] = []
    for line in lines[start + 1 :]:
        value = plain(line).strip()
        if value.startswith("对标真题") or value in MAIN_SECTION_TITLES:
            break
        for term in span_terms(line, "dot"):
            if term not in terms:
                terms.append(term)
    return terms


def deep_glossary_background_terms(lines: list[str]) -> list[str]:
    start = glossary_prompt_index(lines)
    if start is None:
        return []
    terms: list[str] = []
    for line in lines[:start]:
        for class_name in ("yuedu-bg-word", "dot"):
            for term in span_terms(line, class_name):
                if term not in terms:
                    terms.append(term)
    for term in deep_glossary_question_terms(lines):
        if term not in terms:
            terms.append(term)
    return terms


def wrap_plain_term_once(line: str, term: str) -> tuple[str, bool]:
    if not term or f'<span class="yuedu-bg-word">{term}</span>' in line:
        return line, False
    parts = re.split(r"(<[^>]+>)", line)
    in_bg = False
    for index, part in enumerate(parts):
        if not part:
            continue
        if part.startswith("<"):
            if "yuedu-bg-word" in part and not part.startswith("</"):
                in_bg = True
            elif part.startswith("</span") and in_bg:
                in_bg = False
            continue
        if in_bg:
            continue
        pos = part.find(term)
        if pos < 0:
            continue
        parts[index] = f'{part[:pos]}<span class="yuedu-bg-word">{term}</span>{part[pos + len(term):]}'
        return "".join(parts), True
    return line, False


def add_glossary_prompt_backgrounds(line: str, remaining_terms: list[str]) -> str:
    for term in list(remaining_terms):
        line, matched = wrap_plain_term_once(line, term)
        if matched:
            remaining_terms.remove(term)
    return line


def cover_volume_label(volume: str) -> str:
    match = re.match(r"^(专题[一二三四五六七八九十]+)(.+)$", volume)
    if match:
        return f"{match.group(1)} {match.group(2)}"
    return volume


def cover_piece_range(start: int, end: int) -> str:
    start_text = ARABIC_NUMBERS.get(start, str(start))
    end_text = ARABIC_NUMBERS.get(end, str(end))
    return f"篇目{start_text}至{end_text}"


DEEP_TEXT_GLOSSARY_TITLE_RE = re.compile(r"^(?:\d+[\.．、]\s*)?(?:解释下列加点的字词|解释文中加底色的字词)")


def is_deep_text_glossary_title(line: str) -> bool:
    return bool(DEEP_TEXT_GLOSSARY_TITLE_RE.match(plain(line).strip()))


def format_deep_text_glossary_body_lines(lines: list[str]) -> list[str]:
    result: list[str] = []
    index = 0
    before_first_glossary_prompt = True
    remaining_background_terms = deep_glossary_background_terms(lines)
    remaining_prompt_terms = deep_glossary_question_terms(lines)
    while index < len(lines):
        line = lines[index]
        next_line = lines[index + 1] if index + 1 < len(lines) else ""
        if "<img" in line and not plain(line).strip() and is_deep_text_glossary_title(next_line):
            line = f"{line}　　　{next_line}"
            index += 2
        else:
            index += 1
        if before_first_glossary_prompt and is_deep_text_glossary_title(line):
            before_first_glossary_prompt = False
        if before_first_glossary_prompt:
            line = dot_terms_to_background(line)
            line = add_glossary_prompt_backgrounds(line, remaining_background_terms)
            line = add_glossary_prompt_backgrounds(line, remaining_prompt_terms)
        if "<img" in line and "解释文中加底色的字词" in plain(line):
            line = re.sub(r"(<img\b[^>]*>\s*)", r"\1　　　　", line, count=1)
        line = re.sub(r"［(\d+)］", r"<sup>[\1]</sup>", line)
        line = highlight_sup_target_words(line)
        long_answer_match: re.Match[str] | None = None
        for match in re.finditer(r"（(<span class=\"answer-fill\">.*?</span>)）", line):
            answer_html = match.group(1)
            answer_text = re.sub(r"^[\[［]\d+[\]］]\s*", "", plain(answer_html)).strip()
            if wenyan.text_units(answer_text) > 14.0:
                long_answer_match = match
                break
        if long_answer_match is not None:
            split_at = long_answer_match.start() + 1
            before = line[:split_at].rstrip()
            after = line[split_at:].lstrip()
            if before:
                result.append(before)
            if after:
                result.append(after)
            continue
        result.append(line)
    return result


def _kept_match_char(ch: str) -> str:
    if ch.isspace() or ch in {"\u3000", "\u00a0", "\r", "\n", "\t", "．", "•", "·", "\u0323"}:
        return ""
    return ch


def _normalized_with_raw_map(text: str) -> tuple[str, list[int]]:
    chars: list[str] = []
    raw_indexes: list[int] = []
    for index, ch in enumerate(text):
        kept = _kept_match_char(ch)
        if not kept:
            continue
        chars.append(kept)
        raw_indexes.append(index)
    return "".join(chars), raw_indexes


def _raw_to_norm_starts(text: str) -> list[int]:
    starts: list[int] = [0] * (len(text) + 1)
    norm_index = 0
    for index, ch in enumerate(text):
        starts[index] = norm_index
        if _kept_match_char(ch):
            norm_index += 1
        starts[index + 1] = norm_index
    return starts


def _merge_ranges(ranges: list[tuple[int, int]]) -> list[tuple[int, int]]:
    merged: list[tuple[int, int]] = []
    for start, end in sorted(ranges):
        if end <= start:
            continue
        if not merged or start > merged[-1][1]:
            merged.append((start, end))
        else:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
    return merged


def _range_has_overlap(start: int, end: int, ranges: list[tuple[int, int]]) -> bool:
    return any(start < range_end and range_start < end for range_start, range_end in ranges)


def _run_has_gray_highlight(run: ET.Element) -> bool:
    for highlight in run.findall(f".//{{{wenyan.A_NS}}}highlight"):
        srgb = highlight.find(f"{{{wenyan.A_NS}}}srgbClr")
        if srgb is not None and (srgb.attrib.get("val") or "").upper() == "D9D9D9":
            return True
    return False


def _existing_gray_highlight_ranges(sp: ET.Element, limit: int) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []
    pos = 0
    for run in sp.findall(f".//{{{wenyan.A_NS}}}r"):
        t_el = run.find(f"{{{wenyan.A_NS}}}t")
        text = t_el.text if t_el is not None and t_el.text else ""
        text_len = len(text)
        if text_len and pos < limit and _run_has_gray_highlight(run):
            ranges.append((pos, min(pos + text_len, limit)))
        pos += text_len
    return _merge_ranges(ranges)


def _shape_name(sp: ET.Element) -> str:
    c_nv_pr = sp.find(f".//{{{wenyan.P_NS}}}cNvPr")
    return c_nv_pr.attrib.get("name", "") if c_nv_pr is not None else ""


def _shape_plain_text(sp: ET.Element) -> str:
    return "".join(t.text or "" for t in sp.findall(f".//{{{wenyan.A_NS}}}t"))


def _is_answer_overlay_shape(sp: ET.Element) -> bool:
    return "answer" in _shape_name(sp).lower()


def _reading_boundary_index(text: str) -> int:
    candidates = [
        text.find(marker)
        for marker in (
            "深研文本",
            "解释下列加点",
            "解释文中加底色",
            "对标真题",
            "教考衔接",
            "联读拓展",
            "文化常识",
            "参考译文",
        )
        if text.find(marker) >= 0
    ]
    return min(candidates) if candidates else len(text)


def _deep_reading_highlight_source(piece: wenyan.Piece) -> tuple[str, list[tuple[int, int]]]:
    body_sections = sanbu.split_sanbu_piece_lines(piece)
    body_main = split_yuedu_main_sections(body_sections.get("main", []), answer=False)
    lines = body_main.get("深研文本", [])
    prompt_index = glossary_prompt_index(lines)
    if prompt_index is None:
        return "", []

    source_parts: list[str] = []
    ranges: list[tuple[int, int]] = []
    raw_pos = 0
    for line in lines[:prompt_index]:
        for inline in wenyan.parse_inline_runs(line):
            if inline.tif_name:
                continue
            text = inline.text or ""
            if not text:
                continue
            start = raw_pos
            source_parts.append(text)
            raw_pos += len(text)
            if inline.highlight == "D9D9D9" or inline.emphasis_dot:
                ranges.append((start, raw_pos))
        source_parts.append("\n")
        raw_pos += 1
    return "".join(source_parts), _merge_ranges(ranges)


def _deep_reading_highlight_specs(
    body_pieces: list[wenyan.Piece],
) -> dict[str, tuple[str, str, list[tuple[int, int]], list[int]]]:
    specs: dict[str, tuple[str, str, list[tuple[int, int]], list[int]]] = {}
    for piece in body_pieces:
        source_text, raw_ranges = _deep_reading_highlight_source(piece)
        if not source_text or not raw_ranges:
            continue
        norm_source, _source_norm_to_raw = _normalized_with_raw_map(source_text)
        raw_to_norm = _raw_to_norm_starts(source_text)
        norm_ranges = _merge_ranges(
            [
                (raw_to_norm[start], raw_to_norm[end])
                for start, end in raw_ranges
                if raw_to_norm[end] > raw_to_norm[start]
            ]
        )
        if norm_source and norm_ranges:
            specs[title_key(piece.title)] = (piece.title, norm_source, norm_ranges, _source_norm_to_raw)
    return specs


def _best_piece_spec_for_shape(
    sp: ET.Element,
    specs: dict[str, tuple[str, str, list[tuple[int, int]], list[int]]],
) -> tuple[str, str, list[tuple[int, int]], list[int]] | None:
    name_key = title_key(_shape_name(sp))
    matches = [
        spec
        for key, spec in specs.items()
        if key and key in name_key
    ]
    if not matches:
        return None
    return max(matches, key=lambda item: len(title_key(item[0])))


def _ensure_gray_highlight(r_pr: ET.Element) -> bool:
    for highlight in r_pr.findall(f"{{{wenyan.A_NS}}}highlight"):
        srgb = highlight.find(f"{{{wenyan.A_NS}}}srgbClr")
        if srgb is not None and (srgb.attrib.get("val") or "").upper() == "D9D9D9":
            return False
    highlight = ET.Element(wenyan.qn(wenyan.A_NS, "highlight"))
    ET.SubElement(highlight, wenyan.qn(wenyan.A_NS, "srgbClr"), {"val": "D9D9D9"})
    insert_at = 0
    for index, child in enumerate(list(r_pr)):
        if child.tag.endswith("solidFill"):
            insert_at = index + 1
            break
        if child.tag.endswith("latin"):
            insert_at = index
            break
    r_pr.insert(insert_at, highlight)
    return True


def _clone_run_with_text(run: ET.Element, text: str, *, highlight: bool) -> ET.Element:
    new_run = copy.deepcopy(run)
    t_el = new_run.find(f"{{{wenyan.A_NS}}}t")
    if t_el is None:
        t_el = ET.SubElement(new_run, wenyan.qn(wenyan.A_NS, "t"))
    t_el.text = text
    if highlight:
        r_pr = new_run.find(f"{{{wenyan.A_NS}}}rPr")
        if r_pr is None:
            r_pr = ET.Element(wenyan.qn(wenyan.A_NS, "rPr"))
            new_run.insert(0, r_pr)
        _ensure_gray_highlight(r_pr)
    return new_run


def _apply_highlight_ranges_to_shape(sp: ET.Element, ranges: list[tuple[int, int]]) -> int:
    ranges = _merge_ranges(ranges)
    if not ranges:
        return 0
    changed_runs = 0
    global_pos = 0
    for paragraph in sp.findall(f".//{{{wenyan.A_NS}}}p"):
        new_children: list[ET.Element] = []
        paragraph_changed = False
        for child in list(paragraph):
            if child.tag != wenyan.qn(wenyan.A_NS, "r"):
                new_children.append(child)
                continue
            t_el = child.find(f"{{{wenyan.A_NS}}}t")
            text = t_el.text if t_el is not None and t_el.text else ""
            if not text:
                new_children.append(child)
                continue
            run_start = global_pos
            run_end = global_pos + len(text)
            if not _range_has_overlap(run_start, run_end, ranges):
                new_children.append(child)
                global_pos = run_end
                continue

            split_points = {0, len(text)}
            for start, end in ranges:
                if run_start < end and start < run_end:
                    split_points.add(max(0, start - run_start))
                    split_points.add(min(len(text), end - run_start))
            points = sorted(split_points)
            for left, right in zip(points, points[1:]):
                if right <= left:
                    continue
                part_text = text[left:right]
                part_start = run_start + left
                part_end = run_start + right
                should_highlight = _range_has_overlap(part_start, part_end, ranges)
                new_children.append(_clone_run_with_text(child, part_text, highlight=should_highlight))
            paragraph_changed = True
            changed_runs += 1
            global_pos = run_end
        if paragraph_changed:
            paragraph[:] = new_children
        else:
            global_pos += 0
    return changed_runs


def _missing_highlight_ranges_for_shape(
    sp: ET.Element,
    spec: tuple[str, str, list[tuple[int, int]], list[int]],
) -> list[tuple[int, int]]:
    text = _shape_plain_text(sp)
    if not text.strip():
        return []
    boundary = _reading_boundary_index(text)
    reading_text = text[:boundary]
    norm_shape, shape_norm_to_raw = _normalized_with_raw_map(reading_text)
    if len(norm_shape) < 2:
        return []

    _piece_title, norm_source, norm_ranges, _source_norm_to_raw = spec
    source_start = norm_source.find(norm_shape)
    if source_start < 0:
        return []
    source_end = source_start + len(norm_shape)
    target_ranges: list[tuple[int, int]] = []
    existing_highlight_ranges = _existing_gray_highlight_ranges(sp, boundary)
    for start, end in norm_ranges:
        if start < source_end and source_start < end:
            local_start = max(start, source_start) - source_start
            local_end = min(end, source_end) - source_start
            if local_end <= local_start:
                continue
            raw_start = shape_norm_to_raw[local_start]
            raw_end = shape_norm_to_raw[local_end - 1] + 1
            if not _range_has_overlap(raw_start, raw_end, existing_highlight_ranges):
                target_ranges.append((raw_start, raw_end))
    return _merge_ranges(target_ranges)


def patch_wenyan_yuedu_highlights_inplace(
    *,
    pptx_path: Path,
    body_pieces: list[wenyan.Piece],
    dry_run: bool = False,
) -> dict[str, int | str]:
    """Patch missing gray backgrounds in existing PPTX without regenerating slides."""

    specs = _deep_reading_highlight_specs(body_pieces)
    if not specs:
        return {"pptx": str(pptx_path), "slides": 0, "shapes": 0, "runs": 0}

    changed_slides = 0
    changed_shapes = 0
    changed_runs = 0
    with zipfile.ZipFile(pptx_path, "r") as zin:
        updates: dict[str, bytes] = {}
        slide_names = sorted(
            [name for name in zin.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", name)],
            key=lambda name: int(re.search(r"(\d+)", name).group(1)),
        )
        for slide_name in slide_names:
            root = ET.fromstring(zin.read(slide_name))
            slide_shape_changes = 0
            for sp in root.findall(f".//{{{wenyan.P_NS}}}sp"):
                if _is_answer_overlay_shape(sp):
                    continue
                spec = _best_piece_spec_for_shape(sp, specs)
                if spec is None:
                    continue
                ranges = _missing_highlight_ranges_for_shape(sp, spec)
                if not ranges:
                    continue
                run_count = _apply_highlight_ranges_to_shape(sp, ranges)
                if run_count:
                    slide_shape_changes += 1
                    changed_shapes += 1
                    changed_runs += run_count
            if slide_shape_changes:
                changed_slides += 1
                updates[slide_name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

        if updates and not dry_run:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                temp_path = Path(tmp.name)
            try:
                with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        data = updates.get(item.filename)
                        if data is None:
                            data = zin.read(item.filename)
                        zout.writestr(item, data)
                shutil.move(str(temp_path), pptx_path)
            finally:
                if temp_path.exists():
                    temp_path.unlink()

    return {
        "pptx": str(pptx_path),
        "slides": changed_slides,
        "shapes": changed_shapes,
        "runs": changed_runs,
    }


def trim_long_answer_analysis_spacers(lines: list[wenyan.StyledLine]) -> list[wenyan.StyledLine]:
    result: list[wenyan.StyledLine] = []
    for line in lines:
        if (
            (line.role or "") == "answer_spacer"
            and result
            and (result[-1].role or "") == "long_answer"
        ):
            continue
        result.append(line)
    return result


def attach_key_ability_to_previous_answer_chunk(
    chunks: list[list[wenyan.StyledLine]],
    *,
    max_units: int = 12,
) -> list[list[wenyan.StyledLine]]:
    result: list[list[wenyan.StyledLine]] = []
    for chunk in chunks:
        current = list(chunk)
        if (
            result
            and current
            and plain(current[0].text).strip().startswith("【关键能力】")
            and result[-1]
            and (result[-1][-1].role or "") == "long_answer"
            and chunk_cost(result[-1] + [current[0]]) <= max_units
        ):
            result[-1].append(current.pop(0))
        if current:
            result.append(current)
    return result


def strip_deep_text_answer_labels(lines: list[str] | None) -> list[str] | None:
    if not lines:
        return lines
    label_re = re.compile(r"(?:［\d+］|\[\d+\]|[（(]\d+[）)])\s*")
    return [label_re.sub("", line) for line in lines]


def split_numbered_answer_line(line: str) -> tuple[int, str] | None:
    match = re.match(r"^\s*(\d+)[\.、]\s*(.*?)\s*$", plain(line))
    if not match:
        return None
    return int(match.group(1)), match.group(2).strip()


def split_circled_answer_values(line: str) -> list[str]:
    text = plain(line)
    if "】" in text:
        text = text.split("】", 1)[1]
    matches = list(re.finditer(rf"([{re.escape(wenyan.CIRCLED_NUMBERS)}])", text))
    values: list[str] = []
    for index, match in enumerate(matches):
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        value = text[start:end].strip()
        if value:
            values.append(value)
    return values


def parse_vocab_answer_map(answer_lines: list[str] | None) -> dict[int, tuple[str, list[str]]]:
    result: dict[int, tuple[str, list[str]]] = {}
    if not answer_lines:
        return result
    current_num: int | None = None
    for line in answer_lines:
        numbered = split_numbered_answer_line(line)
        if numbered:
            current_num, answer = numbered
            result[current_num] = (answer, [])
            continue
        if current_num is not None and "课内链接" in plain(line):
            main, _old_links = result.get(current_num, ("", []))
            result[current_num] = (main, split_circled_answer_values(line))
    return result


def fill_first_plain_blank(line: str, answer: str) -> str:
    if not answer:
        return line
    return re.sub(r"[_＿]{2,}", wenyan.answer_span(answer, underline=True, pad_underline=True), line, count=1)


def fill_first_vocab_blank(line: str, answer: str) -> tuple[str, bool]:
    if not answer:
        return line, False
    filled_lines, matched = wenyan.fill_lines_with_answer_values([line], [answer])
    if matched:
        return filled_lines[0], True
    filled = fill_first_plain_blank(line, answer)
    return filled, filled != line


def split_vocab_numbered_prompt(line: str) -> list[str]:
    """Separate a numeric glossary prompt from its right-side answer slot."""

    if "<span" not in line or "answer-fill" not in line:
        return [line]
    dot_terms = re.findall(r'<span\s+class="dot"[^>]*>(.*?)</span>', line)
    if not dot_terms:
        return [line]
    target = plain(dot_terms[-1]).strip()
    if not target:
        return [line]

    answer_re = re.compile(
        rf"({re.escape(target)}[：:]\s*<span\s+class=\"answer-fill[^\"]*\">.*?</span>)\s*$",
        re.S,
    )
    match = answer_re.search(line)
    if not match:
        return [line]

    left = line[: match.start()].rstrip()
    right = match.group(1).strip()
    left_units = wenyan.text_units(plain(left))
    right_units = wenyan.text_units(plain(right))
    if left_units + right_units + 4.0 <= 31.5:
        gap_units = max(4.0, 22.0 - left_units)
        return [f"{left}{'　' * int(math.ceil(gap_units))}{right}"]
    return [left, f"{'　' * 12}{right}"]


def is_vocab_link_icon_line(line: str) -> bool:
    return "<img" in line and "B5.tif" in line


def build_yuedu_vocab_styled_section(
    lines: list[str],
    answer_lines: list[str] | None,
) -> tuple[list[wenyan.StyledLine], bool]:
    answer_map = parse_vocab_answer_map(answer_lines)
    styled: list[wenyan.StyledLine] = []
    current_num: int | None = None
    current_link_index = 0
    pending_main_answer = ""
    matched = False

    for raw_line in lines:
        if is_vocab_link_icon_line(raw_line):
            styled.append(wenyan.StyledLine(raw_line, "vocab_link_icon"))
            continue
        value = plain(raw_line).strip()
        if not value or value in {"教考衔接", "重点字词"}:
            continue
        numbered = split_numbered_answer_line(raw_line)
        if numbered:
            current_num = numbered[0]
            current_link_index = 0
            main_answer, _links = answer_map.get(current_num, ("", []))
            filled_line, filled = fill_first_vocab_blank(raw_line, main_answer)
            pending_main_answer = "" if filled else main_answer
            styled.extend(wenyan.styled_body_lines(split_vocab_numbered_prompt(filled_line), lines[0] if lines else ""))
            matched = matched or filled
            continue
        is_circled_link_line = bool(re.match(rf"^[{re.escape(wenyan.CIRCLED_NUMBERS)}]", value))
        if (
            current_num is not None
            and pending_main_answer
            and not is_circled_link_line
            and wenyan.line_has_fillable_blank(raw_line)
        ):
            filled_line, filled = fill_first_vocab_blank(raw_line, pending_main_answer)
            if filled:
                pending_main_answer = ""
                styled.extend(wenyan.styled_body_lines([filled_line], lines[0] if lines else ""))
                matched = True
                continue
        if current_num is not None and is_circled_link_line:
            _main_answer, link_answers = answer_map.get(current_num, ("", []))
            answer = link_answers[current_link_index] if current_link_index < len(link_answers) else ""
            current_link_index += 1
            filled_line, filled = fill_first_vocab_blank(raw_line, answer)
            styled.extend(wenyan.styled_body_lines([filled_line], lines[0] if lines else ""))
            matched = matched or filled
            continue
        styled.extend(wenyan.styled_body_lines([raw_line], lines[0] if lines else ""))

    return styled, matched


def trim_answer_reference_material_tail(answer_lines: list[str] | None) -> list[str] | None:
    if not answer_lines:
        return answer_lines
    trimmed: list[str] = []
    for line in answer_lines:
        if is_material_label_text(line):
            break
        trimmed.append(line)
    return trimmed


def is_blank_only_answer_line(line: str) -> bool:
    value = plain(line).strip()
    return bool(value) and bool(re.fullmatch(r"[_＿\s]+", value))


def open_answer_line_count(lines: list[str]) -> int:
    return sum(1 for line in lines if is_blank_only_answer_line(line))


def is_open_answer_group(lines: list[str]) -> bool:
    return open_answer_line_count(lines) > 0


def remove_open_answer_blank_lines(lines: list[str]) -> list[str]:
    cleaned = [line for line in lines if not is_blank_only_answer_line(line)]
    return cleaned or lines


def split_item_answer_analysis(item: wenyan.AnswerItem) -> tuple[list[str], list[str]]:
    answers: list[str] = []
    analyses = list(item.analyses)
    for value in item.answers:
        if ANALYSIS_LABEL_RE.match(plain(value).strip()):
            analyses.append(value)
        else:
            answers.append(value)
    return answers, analyses


def extract_numbered_answer_block(answer_lines: list[str] | None, num: int) -> list[str]:
    if not answer_lines:
        return []
    block: list[str] = []
    collecting = False
    for line in answer_lines:
        numbered = split_numbered_answer_line(line)
        if numbered:
            if collecting:
                break
            if numbered[0] == num:
                collecting = True
                if numbered[1]:
                    block.append(numbered[1])
            continue
        if collecting:
            block.append(line)
    return block


def split_translation_answer_block(block: list[str]) -> tuple[list[str], list[str]]:
    answers: list[str] = []
    analyses: list[str] = []
    in_analysis = False
    for line in block:
        value = plain(line).strip()
        if not value:
            continue
        if ANALYSIS_LABEL_RE.match(value):
            in_analysis = True
            analyses.append(line)
            continue
        if in_analysis:
            analyses.append(line)
            continue
        if re.match(r"^\s*[（(]\d+[）)]", value):
            answers.append(strip_subanswer_label(value))
        else:
            answers.append(value)
    return answers, analyses


def is_translation_group(group_lines: list[str]) -> bool:
    return bool(group_lines and "翻译成现代汉语" in plain(group_lines[0]))


def strip_subanswer_label(text: str) -> str:
    return re.sub(r"^[（(]\d+[）)]\s*", "", text).strip()


def yuedu_fixed_answer_line(answer: str = "") -> str:
    answer = answer.strip()
    leading = "　" if answer else ""
    answer_units = wenyan.text_units(f"{leading}{answer}")
    padding_units = max(
        int(YUEDU_LONG_ANSWER_TRAILING_UNITS),
        int(math.ceil(YUEDU_LONG_ANSWER_UNDERLINE_UNITS - answer_units)),
    )
    # Long-answer blanks must live in the same text box as the question.  The
    # visible answer is animated as usual, while the underlined NBSP padding
    # keeps the answer area editable like a typed underline instead of a shape.
    padding = "\u00a0" * (padding_units * 2)
    return wenyan.answer_span(f"{leading}{answer}{padding}", underline=True)


def yuedu_split_long_answer_text(answer: str, max_units: float = YUEDU_LONG_ANSWER_TEXT_UNITS) -> list[str]:
    text = re.sub(r"\s+", "", answer).strip()
    if not text:
        return []
    pieces: list[str] = []
    break_chars = "，。！？；、）)]"
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
            if unit_count >= max_units - YUEDU_LONG_ANSWER_TRAILING_UNITS
        ]
        cut = candidates[-1] if candidates else hard_cut
        if cut <= 0:
            cut = 1
        pieces.append(text[:cut].strip())
        text = text[cut:].strip()
    if len(pieces) >= 2 and wenyan.text_units(pieces[-1]) <= YUEDU_LONG_ANSWER_TRAILING_UNITS:
        previous_units = wenyan.text_units(pieces[-2])
        tail_units = wenyan.text_units(pieces[-1])
        if previous_units + tail_units <= YUEDU_LONG_ANSWER_UNDERLINE_UNITS:
            pieces[-2] = f"{pieces[-2]}{pieces[-1]}"
            pieces.pop()
    return [piece for piece in pieces if piece]


def yuedu_centered_long_answer_lines(
    answer: str,
    *,
    max_units: float = YUEDU_LONG_ANSWER_TEXT_UNITS,
) -> list[wenyan.StyledLine]:
    pieces = yuedu_split_long_answer_text(answer, max_units=max_units)
    if not pieces:
        pieces = [""]
    return [
        wenyan.StyledLine(yuedu_fixed_answer_line(piece), "long_answer")
        for piece in pieces
    ]


def build_translation_group_styled(
    group_lines: list[str],
    item: wenyan.AnswerItem,
    section_title: str,
    answer_block: list[str] | None = None,
) -> list[wenyan.StyledLine]:
    answers, analyses = (
        split_translation_answer_block(answer_block)
        if answer_block
        else split_item_answer_analysis(item)
    )
    styled: list[wenyan.StyledLine] = []
    styled.extend(wenyan.styled_body_lines([group_lines[0]], section_title))
    answer_index = 0
    for line in group_lines[1:]:
        if is_blank_only_answer_line(line):
            continue
        styled.extend(wenyan.styled_body_lines([line], section_title))
        if re.match(r"^\s*[（(]\d+[）)]", plain(line)):
            answer = strip_subanswer_label(answers[answer_index]) if answer_index < len(answers) else ""
            if answer:
                styled.extend(yuedu_centered_long_answer_lines(answer))
            answer_index += 1
    if analyses:
        styled.append(wenyan.StyledLine("", "answer_spacer"))
        for analysis in analyses:
            styled.append(wenyan.StyledLine(analysis, "analysis"))
    return styled


def build_unnumbered_long_answer_section(
    lines: list[str],
    answer_lines: list[str] | None,
    section_title: str,
) -> tuple[list[wenyan.StyledLine], bool]:
    if not answer_lines or plain(section_title).strip() != "联读拓展":
        return [], False
    answer_parts: list[str] = []
    analyses: list[str] = []
    for line in answer_lines:
        value = plain(line).strip()
        if not value or value == "联读拓展":
            continue
        if ANALYSIS_LABEL_RE.match(value):
            analyses.append(line)
        else:
            answer_parts.append(line)
    if not answer_parts:
        return [], False
    body_lines = [line for line in lines if not is_blank_only_answer_line(line)]
    styled = wenyan.styled_body_lines(body_lines, section_title)
    long_answer_count = 0
    for answer in answer_parts:
        for answer_line in yuedu_centered_long_answer_lines(plain(answer)):
            if long_answer_count == 5:
                styled.append(wenyan.StyledLine("", "page_break"))
                answer_line = wenyan.StyledLine(
                    answer_line.text,
                    answer_line.role,
                    {"keep_new_chunk": True},
                )
            styled.append(answer_line)
            long_answer_count += 1
    if analyses:
        styled.append(wenyan.StyledLine("", "answer_spacer"))
        for analysis in analyses:
            styled.append(wenyan.StyledLine(analysis, "analysis"))
    return styled, True


def build_yuedu_styled_section(
    lines: list[str],
    answer_lines: list[str] | None,
) -> tuple[list[wenyan.StyledLine], bool]:
    if lines and plain(lines[0]).strip() == "教考衔接":
        return build_yuedu_vocab_styled_section(lines, answer_lines)

    section_title = lines[0] if lines else ""
    answer_items = wenyan.extract_answer_items(answer_lines)
    leading, groups = wenyan.split_body_numbered_groups(lines)
    if not (answer_items and groups):
        unnumbered_styled, unnumbered_matched = build_unnumbered_long_answer_section(lines, answer_lines, section_title)
        if unnumbered_matched:
            return unnumbered_styled, True
        return sanbu.build_sanbu_styled_section(lines, answer_lines)

    styled: list[wenyan.StyledLine] = []
    preamble_lines, preamble_filled = wenyan.fill_lines_with_answer_values(
        leading,
        wenyan.extract_preamble_answers(answer_lines),
    )
    styled.extend(wenyan.styled_body_lines(preamble_lines, section_title))
    matched_any = preamble_filled

    for num, group_lines in groups:
        item = answer_items.get(num)
        if item is not None and is_translation_group(group_lines):
            styled.extend(
                build_translation_group_styled(
                    group_lines,
                    item,
                    section_title,
                    extract_numbered_answer_block(answer_lines, num),
                )
            )
            matched_any = True
            continue
        if item is not None and is_open_answer_group(group_lines):
            question_lines = remove_open_answer_blank_lines(group_lines)
            styled.extend(wenyan.styled_body_lines(question_lines, section_title))
            answers, analyses = split_item_answer_analysis(item)
            long_answer_lines = yuedu_centered_long_answer_lines("；".join(answers))
            styled.extend(long_answer_lines)
            if analyses:
                styled.append(wenyan.StyledLine("", "answer_spacer"))
            for analysis in analyses:
                styled.append(wenyan.StyledLine(analysis, "analysis"))
            matched_any = True
            continue

        partial_styled, partial_matched = sanbu.build_sanbu_styled_section(group_lines, answer_lines)
        styled.extend(partial_styled)
        matched_any = matched_any or partial_matched
    return styled, matched_any


def style_yuedu_styled_lines(lines: list[wenyan.StyledLine]) -> list[wenyan.StyledLine]:
    styled: list[wenyan.StyledLine] = []
    in_deep_glossary_reading = False
    in_material_reading = False
    for line in lines:
        role = line.role
        text = line.text
        if is_deep_text_glossary_title(text):
            in_deep_glossary_reading = True
            in_material_reading = False
        elif is_material_label_text(text):
            role = "material_label"
            in_material_reading = True
        elif (in_deep_glossary_reading or in_material_reading) and role in {"body", "reading", "extended_reading"}:
            role = "deep_material_reading"
        if role in {"analysis", "analysis_block", "segmentation_analysis"}:
            text = style_analysis_label_runs(text)
        styled.append(wenyan.StyledLine(text, role, line.meta))
    return styled


def split_overlong_glossary_answer_lines(lines: list[wenyan.StyledLine]) -> list[wenyan.StyledLine]:
    result: list[wenyan.StyledLine] = []
    answer_re = re.compile(r"（(<span class=\"answer-fill\">.*?</span>)）")
    for line in lines:
        text = line.text
        split_at: int | None = None
        for match in answer_re.finditer(text):
            answer_text = re.sub(r"^[\[［]\d+[\]］]\s*", "", plain(match.group(1))).strip()
            if wenyan.text_units(answer_text) > 14.0:
                split_at = match.start()
                break
        if split_at is None:
            result.append(line)
            continue
        head = text[:split_at].rstrip()
        tail = text[split_at:].lstrip()
        if head:
            result.append(wenyan.StyledLine(head, line.role, line.meta))
        if tail:
            result.append(wenyan.StyledLine(tail, line.role, line.meta))
    return result


def freeze_analysis_roles(chunks: list[list[wenyan.StyledLine]]) -> list[list[wenyan.StyledLine]]:
    result: list[list[wenyan.StyledLine]] = []
    for chunk in chunks:
        converted: list[wenyan.StyledLine] = []
        for line in chunk:
            role = "yuedu_analysis" if line.role in {"analysis", "analysis_block", "segmentation_analysis"} else line.role
            converted.append(wenyan.StyledLine(line.text, role, line.meta))
        result.append(converted)
    return result


def line_has_answer_content(line: wenyan.StyledLine) -> bool:
    role = line.role or "body"
    if role in {"answer", "analysis", "analysis_block", "segmentation_analysis", "yuedu_analysis", "long_answer"}:
        return True
    return any(run.answer_fill for run in wenyan.parse_inline_runs(line.text))


def chunk_cost(lines: list[wenyan.StyledLine]) -> int:
    return sum(yuedu_line_page_cost(line) for line in lines)


def yuedu_line_page_cost(line: wenyan.StyledLine) -> int:
    if (line.role or "") == "long_answer":
        return 1
    base = getattr(wenyan, "_yuedu_original_line_page_cost", None)
    if base is not None:
        return base(line)
    return wenyan.line_page_cost(line)


def chunk_material_group(group: list[wenyan.StyledLine], max_units: int) -> list[list[wenyan.StyledLine]]:
    chunks: list[list[wenyan.StyledLine]] = []
    current: list[wenyan.StyledLine] = []
    used = 0
    for source_line in group:
        expanded = wenyan.split_oversized_styled_line(source_line)
        for line in expanded:
            cost = yuedu_line_page_cost(line)
            if current and used + cost > max_units:
                chunks.append(current)
                current = []
                used = 0
            current.append(line)
            used += cost
    if current:
        chunks.append(current)
    return chunks


def chunk_deep_text_glossary_styled(
    lines: list[wenyan.StyledLine],
    *,
    max_units: int = 12,
) -> list[list[wenyan.StyledLine]]:
    prefix: list[wenyan.StyledLine] = []
    material_groups: list[list[wenyan.StyledLine]] = []
    current_group: list[wenyan.StyledLine] | None = None

    for line in lines:
        if line.role == "material_label":
            current_group = [line]
            material_groups.append(current_group)
            continue
        if current_group is None:
            prefix.append(line)
        else:
            current_group.append(line)

    if not material_groups:
        return wenyan.chunk_styled_lines(lines, max_units=max_units)

    chunks: list[list[wenyan.StyledLine]] = []
    current = list(prefix)
    for group in material_groups:
        group_cost = chunk_cost(group)
        if group_cost <= max_units:
            if current and chunk_cost(current) + group_cost > max_units:
                chunks.append(current)
                current = []
            current.extend(group)
            continue

        if current:
            chunks.append(current)
            current = []
        chunks.extend(chunk_material_group(group, max_units))

    if current:
        chunks.append(current)
    return chunks


def chunk_vocab_styled_lines(
    lines: list[wenyan.StyledLine],
    *,
    max_units: int = 12,
) -> list[list[wenyan.StyledLine]]:
    prefix: list[wenyan.StyledLine] = []
    groups: list[list[wenyan.StyledLine]] = []
    current_group: list[wenyan.StyledLine] | None = None
    for line in lines:
        if re.match(r"^\s*\d+[\.、]", plain(line.text)):
            current_group = [line]
            groups.append(current_group)
            continue
        if current_group is None:
            prefix.append(line)
        else:
            current_group.append(line)

    if not groups:
        return wenyan.chunk_styled_lines(lines, max_units=max_units)

    chunks: list[list[wenyan.StyledLine]] = []
    current = list(prefix)
    for group in groups:
        cost = chunk_cost(group)
        if cost <= max_units:
            if current and chunk_cost(current) + cost > max_units:
                chunks.append(current)
                current = []
            current.extend(group)
            continue
        if current:
            chunks.append(current)
            current = []
        chunks.extend(wenyan.chunk_styled_lines(group, max_units=max_units))
    if current:
        chunks.append(current)
    return [chunk for chunk in chunks if chunk]


def split_questions_after_answer_content(
    chunks: list[list[wenyan.StyledLine]],
) -> list[list[wenyan.StyledLine]]:
    result: list[list[wenyan.StyledLine]] = []
    for chunk in chunks:
        split_at: int | None = None
        seen_answer = False
        for index, line in enumerate(chunk):
            if index > 0 and seen_answer and wenyan.is_numbered_question_line(line):
                split_at = index
                break
            seen_answer = seen_answer or line_has_answer_content(line)
        if split_at is None:
            result.append(chunk)
            continue
        before = chunk[:split_at]
        after = chunk[split_at:]
        if before:
            result.append(before)
        if after:
            result.append(after)
    result = [chunk for chunk in result if chunk]
    result = merge_yuedu_split_long_answer_chunks(result, max_units=12)
    return sanbu.split_numbered_prompt_after_answer_chunks(result)


def merge_yuedu_split_long_answer_chunks(
    chunks: list[list[wenyan.StyledLine]],
    *,
    max_units: int = 12,
) -> list[list[wenyan.StyledLine]]:
    result: list[list[wenyan.StyledLine]] = []
    for chunk in chunks:
        current = list(chunk)
        if current and (current[0].meta or {}).get("keep_new_chunk"):
            result.append(current)
            continue
        previous_can_take_answer = bool(
            result
            and result[-1]
            and (
                (result[-1][-1].role or "") == "long_answer"
                or (
                    any(sanbu.is_sanbu_numbered_prompt(line) for line in result[-1])
                    and not any(
                        (line.role or "") in {"answer", "long_answer", "answer_spacer", "analysis"}
                        for line in result[-1]
                    )
                )
            )
        )
        if previous_can_take_answer and current and (current[0].role or "") == "long_answer":
            carry: list[wenyan.StyledLine] = []
            while current and (current[0].role or "") == "long_answer":
                carry.append(current.pop(0))
            optional: list[wenyan.StyledLine] = []
            while current and (current[0].role or "") in {"answer_spacer", "analysis"}:
                optional.append(current.pop(0))
            if sanbu.sanbu_chunk_cost(result[-1] + carry + optional) <= max_units + 2:
                result[-1].extend(carry + optional)
            else:
                result[-1].extend(carry)
                current = optional + current
            if current:
                result.append(current)
            continue
        result.append(current)
    return result


def make_yuedu_piece_slide_chunks(body_piece: wenyan.Piece, answer_piece: wenyan.Piece | None) -> list[list[wenyan.StyledLine]]:
    body_sections = sanbu.split_sanbu_piece_lines(body_piece)
    answer_sections = sanbu.split_sanbu_piece_lines(answer_piece)
    body_main = split_yuedu_main_sections(body_sections.get("main", []), answer=False)
    answer_main = split_yuedu_main_sections(answer_sections.get("main", []), answer=True)

    chunks: list[list[wenyan.StyledLine]] = []
    section_order = ["深研文本", "对标真题", "教考衔接", "联读拓展", "文化常识"]
    for section_name in section_order:
        body_lines = body_main.get(section_name, [])
        if not body_lines:
            continue
        answer_lines = answer_main.get(section_name)
        if section_name == "深研文本":
            body_lines = format_deep_text_glossary_body_lines(body_lines)
            answer_lines = strip_deep_text_answer_labels(answer_lines)
        else:
            answer_lines = trim_answer_reference_material_tail(answer_lines)
        styled, _matched = build_yuedu_styled_section(body_lines, answer_lines)
        styled = sanbu.add_slash_answer_labels(styled)
        styled = style_yuedu_styled_lines(styled)
        styled = trim_long_answer_analysis_spacers(styled)
        if section_name == "深研文本" and any(is_material_label_text(line.text) for line in styled):
            section_chunks = chunk_deep_text_glossary_styled(styled, max_units=14)
        elif section_name == "深研文本" and any(is_deep_text_glossary_title(line.text) for line in styled):
            section_chunks = wenyan.chunk_styled_lines(styled, max_units=14)
        elif section_name == "教考衔接":
            section_chunks = chunk_vocab_styled_lines(styled, max_units=12)
        else:
            section_chunks = wenyan.chunk_styled_lines(styled, max_units=12)
        section_chunks = sanbu.repair_orphan_sanbu_question_chunks(section_chunks)
        section_chunks = merge_sparse_intro_chunks(section_chunks)
        section_chunks = split_questions_after_answer_content(section_chunks)
        section_chunks = freeze_analysis_roles(section_chunks)
        section_chunks = attach_key_ability_to_previous_answer_chunk(section_chunks)
        chunks.extend(section_chunks)

    reference_lines = answer_sections.get("reference", [])
    reference_styled: list[wenyan.StyledLine] = []
    has_reference_tif = False
    for line in reference_lines:
        lower = line.lower()
        if any(name.lower() in lower for name in REFERENCE_TIF_NAMES):
            if "参考译文.tif" in lower:
                has_reference_tif = True
                reference_styled.append(wenyan.StyledLine(line, "reference_icon"))
            continue
        value = plain(line).strip()
        if not value or "参考译文" in value:
            continue
        reference_styled.append(wenyan.StyledLine(line, "reference_translation"))
    if reference_styled:
        styled = reference_styled if has_reference_tif else [wenyan.StyledLine("参考译文", "reference_header"), *reference_styled]
        chunks.extend(wenyan.chunk_styled_lines(styled, max_units=9))
    return chunks


def add_yuedu_cover_content(root, volume: str, start: int, end: int) -> None:
    topic = cover_volume_label(volume)
    piece_range = cover_piece_range(start, end)
    for sp in wenyan.text_shapes(root):
        text = wenyan.get_shape_text(sp).strip()
        if "高考文言文阅读加分练" in text:
            wenyan.set_shape_text(sp, "高考文言文阅读加分练")
            wenyan.style_existing_text_shape(sp, font="Songti SC Black", color="000000", bold=True)
            continue
        compact = re.sub(r"\s+", "", text)
        if (
            compact in {volume, topic, f"篇目{start}-{end}", piece_range}
            or re.fullmatch(r"篇目\d+-\d+", compact)
            or re.fullmatch(r"篇目[一二三四五六七八九十]+至[一二三四五六七八九十]+", compact)
            or compact in {f"{volume}篇目{start}-{end}", f"{topic}{piece_range}"}
        ):
            wenyan.set_shape_text(sp, "")
    wenyan.append_textbox(
        root,
        x=7.65,
        y=4.13,
        w=3.55,
        h=0.58,
        lines=[wenyan.StyledLine(topic, "cover_topic")],
        name="文言文阅读加分练封面专题",
    )
    wenyan.append_textbox(
        root,
        x=5.70,
        y=5.13,
        w=4.45,
        h=0.62,
        lines=[wenyan.StyledLine(piece_range, "cover_range")],
        name="文言文阅读加分练封面篇目",
    )
    for sp in wenyan.text_shapes(root):
        name_el = sp.find(f".//{{{wenyan.P_NS}}}cNvPr")
        shape_name = name_el.attrib.get("name", "") if name_el is not None else ""
        if "封面专题" in shape_name:
            for r_pr in sp.findall(f".//{{{wenyan.A_NS}}}rPr"):
                r_pr.set("i", "1")


def add_yuedu_content(
    root,
    rels_root,
    piece: wenyan.Piece,
    chunk: list[wenyan.StyledLine],
    *,
    media_assets: dict[str, bytes],
    tif_search_dirs: list[Path],
    content_box: tuple[float, float, float, float],
    fallback_add_content,
) -> None:
    if not any(line.role == "vocab_link_icon" for line in chunk):
        fallback_add_content(
            root,
            rels_root,
            piece,
            chunk,
            media_assets=media_assets,
            tif_search_dirs=tif_search_dirs,
            content_box=content_box,
        )
        return

    wenyan.clear_text_shapes(root, keep_substrings=("能力",))
    base_x, base_y, base_w, base_h = content_box
    wenyan.append_textbox(
        root,
        x=base_x,
        y=base_y,
        w=base_w,
        h=base_h,
        lines=chunk,
        name=f"{piece.title} vocab content",
        answer_mode="base",
    )
    first_visible_text = next((plain(line.text).strip() for line in chunk if plain(line.text).strip()), "")
    b5_y_offset = 0.32 if first_visible_text.startswith("解释下列加点的字词") else 0.22
    for tif_name, img_x, img_y, img_h in wenyan.collect_tif_positions(x=base_x, y=base_y, w=base_w, lines=chunk):
        tif_path = wenyan.resolve_tif_path(tif_name, tif_search_dirs)
        if tif_path is None:
            continue
        if tif_name.lower() == "b5.tif":
            img_x += 0.70
            img_y += b5_y_offset
        wenyan.append_tif_picture(
            root,
            rels_root,
            media_assets,
            tif_path=tif_path,
            x=img_x,
            y=img_y,
            h=img_h,
            name=f"yuedu inline tif {tif_name}",
        )
    animations = wenyan.append_answer_overlays(
        root,
        x=base_x,
        y=base_y,
        w=base_w,
        h=base_h,
        lines=chunk,
        name=f"{piece.title} vocab content",
    )
    wenyan.set_slide_animations(root, animations)


def add_yuedu_piece_title(root, piece: wenyan.Piece) -> None:
    wenyan.clear_text_shapes(root)
    wenyan.append_piece_title_textbox(root, piece)
    wenyan.append_textbox(
        root,
        x=0.26,
        y=6.42,
        w=4.75,
        h=0.52,
        lines=[wenyan.StyledLine(wenyan.spaced_volume_label(piece.volume), "title_volume")],
        name="文言阅读标题页册别",
    )


def install_yuedu_hooks() -> tuple[object, ...]:
    sanbu.install_sanbu_hooks()
    old_make_piece_slide_chunks = wenyan.make_piece_slide_chunks
    old_add_cover_content = wenyan.add_cover_content
    old_add_piece_title = wenyan.add_piece_title
    old_role_style = wenyan.role_style
    old_is_long_answer_prompt = wenyan.is_long_answer_prompt
    old_add_content = wenyan.add_content
    old_paragraph_properties = wenyan.paragraph_properties
    old_line_page_cost = wenyan.line_page_cost
    old_yuedu_original_line_page_cost = getattr(wenyan, "_yuedu_original_line_page_cost", None)
    old_append_answer_overlays = wenyan.append_answer_overlays
    old_append_tif_overlays = wenyan.append_tif_overlays
    old_sanbu_long_answer_underlines = sanbu.append_sanbu_long_answer_underlines
    old_inline_vertical_shift = wenyan.ANSWER_INLINE_VERTICAL_SHIFT_PT
    old_multiline_vertical_shift = wenyan.ANSWER_MULTILINE_VERTICAL_SHIFT_PT

    def yuedu_role_style(role: str) -> tuple[str, float, str, bool]:
        if role in {"cover_topic", "cover_range"}:
            return "Songti SC Black", 40, "000000", True
        if role == "material_label":
            return "黑体", 24, "000000", True
        if role == "deep_material_reading":
            return "KaiTi", 24, "000000", False
        if role == "yuedu_analysis":
            return "宋体", 24, "FF0000", True
        if role == "vocab_link_icon":
            return "宋体", 24, "000000", False
        return old_role_style(role)

    def yuedu_paragraph_properties(role: str) -> dict[str, str]:
        if role in {"cover_topic", "cover_range"}:
            attrs = old_paragraph_properties("center_heading")
            attrs["algn"] = "ctr"
            return attrs
        if role == "deep_material_reading":
            attrs = old_paragraph_properties("extended_reading")
            attrs["indent"] = str(wenyan.FIRST_LINE_INDENT_EMU)
            attrs["algn"] = "just"
            return attrs
        if role == "yuedu_analysis":
            return old_paragraph_properties("analysis")
        if role == "long_answer":
            attrs = old_paragraph_properties(role)
            attrs["indent"] = "0"
            attrs["marL"] = "0"
            attrs["algn"] = "ctr"
            return attrs
        return old_paragraph_properties(role)

    def yuedu_is_long_answer_prompt(section_title: str, group_lines: list[str]) -> bool:
        if is_choice_question_group(group_lines):
            return False
        return old_is_long_answer_prompt(section_title, group_lines)

    def yuedu_add_content(
        root,
        rels_root,
        piece: wenyan.Piece,
        chunk: list[wenyan.StyledLine],
        *,
        media_assets: dict[str, bytes],
        tif_search_dirs: list[Path],
        content_box: tuple[float, float, float, float],
    ) -> None:
        add_yuedu_content(
            root,
            rels_root,
            piece,
            chunk,
            media_assets=media_assets,
            tif_search_dirs=tif_search_dirs,
            content_box=content_box,
            fallback_add_content=old_add_content,
        )

    def yuedu_append_long_answer_underlines(root, *, x: float, y: float, w: float, lines: list[wenyan.StyledLine]) -> None:
        return None

    def yuedu_append_answer_overlays(
        root,
        *,
        x: float,
        y: float,
        w: float,
        h: float,
        lines: list[wenyan.StyledLine],
        name: str,
    ) -> list[wenyan.AnimationTarget]:
        has_long_answer = any((line.role or "body") == "long_answer" for line in lines)
        choice_analysis_mode = (not has_long_answer) and is_choice_analysis_chunk(lines)
        has_long_plain_analysis = any(
            (line.role or "") in {"analysis", "analysis_block", "segmentation_analysis", "yuedu_analysis"}
            and "<" not in line.text
            and wenyan.text_units(plain(line.text)) > 34
            for line in lines
        )
        if not has_long_answer and not choice_analysis_mode and not has_long_plain_analysis:
            return old_append_answer_overlays(root, x=x, y=y, w=w, h=h, lines=lines, name=name)

        animations: list[wenyan.AnimationTarget] = []
        event_index = 0
        cursor_y = 0.0
        line_index = 0
        overlay_line_factor = 1.32
        choice_body_line_factor = 1.18
        choice_body_capacity_units = 34.8

        def visual_line_count(line: wenyan.StyledLine) -> int:
            role = line.role or "body"
            if role == "long_answer":
                return 1
            if choice_analysis_mode and role in {"body", "reading", "extended_reading"}:
                units = wenyan.text_units(wenyan.strip_inline_tags(line.text))
                return max(1, math.ceil(units / choice_body_capacity_units))
            return max(1, wenyan.estimate_line_count(line.text))

        def overlay_line_step(line: wenyan.StyledLine) -> float:
            role = line.role or "body"
            factor = (
                choice_body_line_factor
                if choice_analysis_mode and role in {"body", "reading", "extended_reading"}
                else overlay_line_factor
            )
            return visual_line_count(line) * wenyan.line_height_for_role(role) * factor

        def event_width(event_lines: list[wenyan.StyledLine]) -> float:
            text_lines = [
                wenyan.strip_inline_tags(line.text).strip()
                for line in event_lines
                if wenyan.strip_inline_tags(line.text).strip()
            ]
            if not text_lines:
                return min(w, 0.35)
            roles = {line.role or "body" for line in event_lines}
            width_pt = wenyan.estimate_answer_overlay_width_pt(text_lines)
            indent_pad = (
                0.72
                if roles
                & {
                    "body",
                    "reading",
                    "extended_reading",
                    "analysis",
                    "analysis_block",
                    "segmentation_analysis",
                    "yuedu_analysis",
                }
                else 0.0
            )
            return max(0.35, min(w, width_pt / 72.0 + indent_pad + 0.08))

        def event_height(event_lines: list[wenyan.StyledLine]) -> float:
            total = 0.0
            for event_line in event_lines:
                total += overlay_line_step(event_line)
            return max(0.25, min(h, total + 0.06))

        def append_line_event(event_lines: list[wenyan.StyledLine], event_y: float, index: int) -> None:
            if not any(wenyan.strip_inline_tags(line.text).strip() for line in event_lines):
                return
            shape_id = wenyan.append_textbox(
                root,
                x=x,
                y=event_y,
                w=event_width(event_lines),
                h=event_height(event_lines),
                lines=event_lines,
                name=f"{name} answer {index + 1}",
                answer_mode="normal",
            )
            animations.append(wenyan.AnimationTarget(shape_id=shape_id, effect="appear", text_target=False))

        def overlay_event_lines(line: wenyan.StyledLine) -> list[wenyan.StyledLine]:
            role = line.role or "body"
            if (
                role in {"analysis", "analysis_block", "segmentation_analysis", "yuedu_analysis"}
                and "<" not in line.text
                and wenyan.text_units(plain(line.text)) > 34
            ):
                return [wenyan.StyledLine(piece, line.role, line.meta) for piece in split_text_by_units(line.text, max_units=30)]
            return [line]

        while line_index < len(lines):
            line = lines[line_index]
            role = line.role or "body"
            if role == "long_answer":
                has_answer = False
                while line_index < len(lines) and (lines[line_index].role or "body") == "long_answer":
                    event_line = lines[line_index]
                    has_answer = has_answer or any(
                        inline.answer_fill for inline in wenyan.parse_inline_runs(event_line.text)
                    )
                    cursor_y += overlay_line_step(event_line)
                    line_index += 1
                if has_answer:
                    shape_id = wenyan.append_textbox(
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
                    animations.append(wenyan.AnimationTarget(shape_id=shape_id, effect="appear", text_target=False))
                    event_index += 1
                continue

            line_h = overlay_line_step(line)
            if role in wenyan.ANIMATED_LINE_ROLES:
                event_lines = overlay_event_lines(line)
                append_line_event(event_lines, y + cursor_y, event_index)
                event_index += 1
                cursor_y += event_height(event_lines)
                line_index += 1
                continue

            inline_answer_count = sum(1 for inline in wenyan.parse_inline_runs(line.text) if inline.answer_fill)
            for _ in range(inline_answer_count):
                shape_id = wenyan.append_textbox(
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
                animations.append(wenyan.AnimationTarget(shape_id=shape_id, effect="appear", text_target=False))
                event_index += 1
            cursor_y += line_h
            line_index += 1
        return animations

    def yuedu_append_tif_overlays(
        root,
        rels_root,
        media_assets: dict[str, bytes],
        *,
        x: float,
        y: float,
        w: float,
        lines: list[wenyan.StyledLine],
        tif_search_dirs: list[Path],
    ) -> None:
        numbered_b10_without_annotation = False
        seen_annotation = False
        for line in lines:
            if (line.role or "") == "annotation":
                seen_annotation = True
            if (
                "B10.tif" in line.text
                and re.match(r"^\s*\d+[\.．、]\s*", plain(line.text).strip())
                and not seen_annotation
            ):
                numbered_b10_without_annotation = True
        for tif_name, img_x, img_y, img_h in wenyan.collect_tif_positions(x=x, y=y, w=w, lines=lines):
            tif_path = wenyan.resolve_tif_path(tif_name, tif_search_dirs)
            if not tif_path:
                continue
            if tif_name.lower() == "参考译文.tif":
                wenyan.append_tif_picture(
                    root,
                    rels_root,
                    media_assets,
                    tif_path=tif_path,
                    x=x,
                    y=y + 0.02,
                    h=0.36,
                    name=f"reference translation tif {tif_name}",
                )
                continue
            if tif_name.lower() == "b10.tif":
                img_x -= 0.50
                img_y += 0.22 if numbered_b10_without_annotation else 0.60
            wenyan.append_tif_picture(
                root,
                rels_root,
                media_assets,
                tif_path=tif_path,
                x=img_x,
                y=img_y,
                h=img_h,
                name=f"fbd tif {tif_name}",
            )

    wenyan.make_piece_slide_chunks = make_yuedu_piece_slide_chunks
    wenyan.add_cover_content = add_yuedu_cover_content
    wenyan.add_piece_title = add_yuedu_piece_title
    wenyan.role_style = yuedu_role_style
    wenyan.is_long_answer_prompt = yuedu_is_long_answer_prompt
    wenyan.paragraph_properties = yuedu_paragraph_properties
    wenyan._yuedu_original_line_page_cost = old_line_page_cost  # type: ignore[attr-defined]
    wenyan.line_page_cost = yuedu_line_page_cost
    wenyan.add_content = yuedu_add_content
    wenyan.append_answer_overlays = yuedu_append_answer_overlays
    wenyan.append_tif_overlays = yuedu_append_tif_overlays
    sanbu.append_sanbu_long_answer_underlines = yuedu_append_long_answer_underlines
    wenyan.ANSWER_INLINE_VERTICAL_SHIFT_PT = 1.0
    wenyan.ANSWER_MULTILINE_VERTICAL_SHIFT_PT = 3.0
    return (
        old_make_piece_slide_chunks,
        old_add_cover_content,
        old_add_piece_title,
        old_role_style,
        old_is_long_answer_prompt,
        old_add_content,
        old_paragraph_properties,
        old_line_page_cost,
        old_yuedu_original_line_page_cost,
        old_append_answer_overlays,
        old_append_tif_overlays,
        old_sanbu_long_answer_underlines,
        old_inline_vertical_shift,
        old_multiline_vertical_shift,
    )


def restore_yuedu_hooks(old_hooks: tuple[object, ...]) -> None:
    (
        old_make_piece_slide_chunks,
        old_add_cover_content,
        old_add_piece_title,
        old_role_style,
        old_is_long_answer_prompt,
        old_add_content,
        old_paragraph_properties,
        old_line_page_cost,
        old_yuedu_original_line_page_cost,
        old_append_answer_overlays,
        old_append_tif_overlays,
        old_sanbu_long_answer_underlines,
        old_inline_vertical_shift,
        old_multiline_vertical_shift,
    ) = old_hooks
    wenyan.make_piece_slide_chunks = old_make_piece_slide_chunks
    wenyan.add_cover_content = old_add_cover_content
    wenyan.add_piece_title = old_add_piece_title
    wenyan.role_style = old_role_style
    wenyan.is_long_answer_prompt = old_is_long_answer_prompt
    wenyan.add_content = old_add_content
    wenyan.paragraph_properties = old_paragraph_properties
    wenyan.line_page_cost = old_line_page_cost
    if old_yuedu_original_line_page_cost is None:
        try:
            delattr(wenyan, "_yuedu_original_line_page_cost")
        except AttributeError:
            pass
    else:
        wenyan._yuedu_original_line_page_cost = old_yuedu_original_line_page_cost  # type: ignore[attr-defined]
    wenyan.append_answer_overlays = old_append_answer_overlays
    wenyan.append_tif_overlays = old_append_tif_overlays
    sanbu.append_sanbu_long_answer_underlines = old_sanbu_long_answer_underlines
    wenyan.ANSWER_INLINE_VERTICAL_SHIFT_PT = old_inline_vertical_shift
    wenyan.ANSWER_MULTILINE_VERTICAL_SHIFT_PT = old_multiline_vertical_shift


@isolated_generation
def generate_wenyan_yuedu_jiafen_pptx(
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
    emit(f"正文 MD：{body_md}")
    emit(f"答案 MD：{answer_md}")
    emit(f"模板 PPTX：{template}")

    answer_pieces = parse_answer_pieces(answer_md)
    body_pieces = parse_body_pieces(body_md, answer_pieces)
    answer_map = {wenyan.piece_key(piece): piece for piece in answer_pieces}
    tif_search_dirs = wenyan.infer_tif_search_dirs(md_run)
    if DEFAULT_SOURCE_DIR.exists() and DEFAULT_SOURCE_DIR not in tif_search_dirs:
        tif_search_dirs.append(DEFAULT_SOURCE_DIR)

    emit(f"正文篇目数：{len(body_pieces)}")
    emit(f"答案篇目数：{len(answer_pieces)}")
    if group_ranges is not None:
        groups = []
        for index, group_range in enumerate(group_ranges, 1):
            volume = str(group_range["volume"])
            start = int(group_range["start"])
            end = int(group_range["end"])
            keys = {tuple(key) for key in group_range["unit_keys"]}  # type: ignore[arg-type]
            selected = [piece for piece in body_pieces if (piece.volume, piece.num) in keys]
            groups.append((custom_yuedu_filename(index, volume, start, end), volume, start, end, selected))
    else:
        base_groups = YUEDU_GROUPS if limit_groups is None else YUEDU_GROUPS[:limit_groups]
        groups = [(*group, None) for group in base_groups]
    output_dir.mkdir(parents=True, exist_ok=True)

    old_hooks = install_yuedu_hooks()
    generated: list[Path] = []
    try:
        for group_index, (file_name, volume, start, end, selected_override) in enumerate(groups, 1):
            group_pieces = selected_override or [piece for piece in body_pieces if piece.volume == volume and start <= piece.num <= end]
            if not group_pieces:
                emit(f"跳过，无篇目：{file_name}")
                continue
            output = output_dir / file_name
            emit(f"[{group_index}/{len(groups)}] 生成：{output.name}")
            wenyan.build_deck(
                template,
                output,
                group_pieces,
                answer_map,
                volume,
                start,
                end,
                tif_search_dirs=tif_search_dirs,
            )
            generated.append(output)
            emit(f"已生成：{output}，篇目：{', '.join(f'{p.num}.{p.title}' for p in group_pieces)}")
    finally:
        restore_yuedu_hooks(old_hooks)

    manifest = {
        "md_run": str(md_run),
        "template": str(template),
        "group_ranges": group_ranges,
        "generated": [str(path) for path in generated],
    }
    (output_dir / "wenyan_yuedu_jiafen_test_generation_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
        newline="\n",
    )
    return generated


def patch_wenyan_yuedu_highlight_batch(
    *,
    md_run: Path,
    pptx_dir: Path,
    dry_run: bool = False,
    backup: bool = True,
    progress: Callable[[str], None] | None = None,
) -> list[dict[str, int | str]]:
    """Patch only missing reading-material gray highlights in existing PPTX files.

    This deliberately does not call the deck generator. It edits the slide XML
    inside the already-created PPTX files and preserves slide order/count.
    """

    def emit(message: str) -> None:
        if progress:
            progress(message)
        else:
            print(message)

    body_md, answer_md = find_md_files(md_run)
    answer_pieces = parse_answer_pieces(answer_md)
    body_pieces = parse_body_pieces(body_md, answer_pieces)
    pptx_files = sorted(path for path in pptx_dir.glob("*.pptx") if not path.name.startswith("~$"))
    if not pptx_files:
        raise FileNotFoundError(f"No PPTX files found in {pptx_dir}")

    if backup and not dry_run:
        backup_dir = pptx_dir / f"_before_highlight_patch_{datetime.now():%Y%m%d_%H%M%S}"
        backup_dir.mkdir(parents=True, exist_ok=True)
        for pptx_path in pptx_files:
            shutil.copy2(pptx_path, backup_dir / pptx_path.name)
        emit(f"Backup created: {backup_dir}")

    results: list[dict[str, int | str]] = []
    for pptx_path in pptx_files:
        result = patch_wenyan_yuedu_highlights_inplace(
            pptx_path=pptx_path,
            body_pieces=body_pieces,
            dry_run=dry_run,
        )
        results.append(result)
        emit(
            f"{pptx_path.name}: slides={result['slides']}, "
            f"shapes={result['shapes']}, runs={result['runs']}"
        )

    manifest = {
        "md_run": str(md_run),
        "pptx_dir": str(pptx_dir),
        "dry_run": dry_run,
        "results": results,
    }
    manifest_name = "wenyan_yuedu_highlight_patch_dry_run.json" if dry_run else "wenyan_yuedu_highlight_patch_manifest.json"
    (pptx_dir / manifest_name).write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
        newline="\n",
    )
    return results


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="测试：文言文阅读加分练 FBD/MD + 模板生成 PPTX。")
    parser.add_argument("--md-run", default="", help="已转换的 MD run 目录；省略时先转换默认 FBD。")
    parser.add_argument("--body-fbd", default=str(DEFAULT_BODY_FBD), help="正文 FBD 路径。")
    parser.add_argument("--answer-fbd", default=str(DEFAULT_ANSWER_FBD), help="答案 FBD 路径。")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="PPTX 模板路径。")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="测试 PPTX 输出目录。")
    parser.add_argument("--limit-groups", type=int, default=1, help="生成前几个分组；默认 1，传 0 表示全部。")
    parser.add_argument("--patch-highlights-inplace", action="store_true", help="只在现有 PPTX 中补阅读材料灰色底色，不重新生成。")
    parser.add_argument("--patch-pptx-dir", default="", help="需要原位修补的现有 PPTX 目录。")
    parser.add_argument("--dry-run", action="store_true", help="与 --patch-highlights-inplace 配合，只统计不写入。")
    parser.add_argument("--no-backup", action="store_true", help="与 --patch-highlights-inplace 配合，不创建备份目录。")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    md_run = Path(args.md_run) if args.md_run else convert_wenyan_yuedu_fbd_pair(
        body_fbd=Path(args.body_fbd),
        answer_fbd=Path(args.answer_fbd),
    )
    if args.patch_highlights_inplace:
        if not args.patch_pptx_dir:
            raise ValueError("--patch-highlights-inplace requires --patch-pptx-dir")
        patch_wenyan_yuedu_highlight_batch(
            md_run=md_run,
            pptx_dir=Path(args.patch_pptx_dir),
            dry_run=args.dry_run,
            backup=not args.no_backup,
        )
        return 0

    limit_groups = None if args.limit_groups == 0 else args.limit_groups
    generate_wenyan_yuedu_jiafen_pptx(
        md_run=md_run,
        template=Path(args.template),
        output_dir=Path(args.output_dir),
        limit_groups=limit_groups,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

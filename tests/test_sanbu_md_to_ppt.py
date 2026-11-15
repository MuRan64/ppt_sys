from __future__ import annotations

import argparse
import html
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Callable


PROJECT_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
if str(Path(__file__).resolve().parent) not in sys.path:
    sys.path.insert(0, str(Path(__file__).resolve().parent))
if str(PROJECT_DIR / "tools") not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR / "tools"))

import test_md_to_ppt as wenyan  # noqa: E402
from generation_state import isolated_generation  # noqa: E402


DEFAULT_MD_RUN = PROJECT_DIR / "out_md" / "三步进阶" / "run_20260708_014621"
DEFAULT_TEMPLATE = PROJECT_DIR / "文言梳理pptx模板.pptx"
DEFAULT_OUTPUT_DIR = PROJECT_DIR / "output" / "test_sanbu_ppt"

SANBU_VOLUMES = {
    "基础篇30篇",
    "提升篇20篇",
    "进阶篇15篇",
}
SANBU_READING_LABELS = {
    "单文本阅读",
    "多文本阅读",
}
SANBU_GROUPS = [
    ("1基础篇30篇 第1-5篇.pptx", "基础篇30篇", 1, 5),
    ("2基础篇30篇 第6-10篇.pptx", "基础篇30篇", 6, 10),
    ("3基础篇30篇 第11-15篇.pptx", "基础篇30篇", 11, 15),
    ("4基础篇30篇 第16-20篇.pptx", "基础篇30篇", 16, 20),
    ("5基础篇30篇 第21-25篇.pptx", "基础篇30篇", 21, 25),
    ("6基础篇30篇 第26-30篇.pptx", "基础篇30篇", 26, 30),
    ("7提升篇20篇 第1-5篇.pptx", "提升篇20篇", 1, 5),
    ("8提升篇20篇 第6-10篇.pptx", "提升篇20篇", 6, 10),
    ("9提升篇20篇 第11-15篇.pptx", "提升篇20篇", 11, 15),
    ("10提升篇20篇 第16-20篇.pptx", "提升篇20篇", 16, 20),
    ("11进阶篇15篇 第1-4篇.pptx", "进阶篇15篇", 1, 4),
    ("12进阶篇15篇 第5-9篇.pptx", "进阶篇15篇", 5, 9),
    ("13进阶篇15篇 第10-15篇.pptx", "进阶篇15篇", 10, 15),
]

MAIN_MARKER = "__SANBU_MAIN__"
REFERENCE_MARKER = "__SANBU_REFERENCE__"
VOCAB_MARKER = "__SANBU_VOCAB__"
L_ICON_MARKER = "__SANBU_L_ICON__"
SANBU_STATIC_DIR = PROJECT_DIR / "tests" / "static" / "三步进阶"
SANBU_SOURCE_DIR = PROJECT_DIR / "三步进阶源文件"
SANBU_ANALYSIS_LABEL_RE = re.compile(r"【(?:解析|关键能力|解题思路|解决思路)】")
SANBU_CIRCLED_NUMBERS = getattr(wenyan, "CIRCLED_NUMBERS", "①②③④⑤⑥⑦⑧⑨⑩")
SANBU_GLOSSARY_PROMPT = "在括号里填写正确的词语释义，并完成"
_SANBU_HOOKS_INSTALLED = False


@dataclass
class SanbuGlossaryEntry:
    label: str
    term: str
    definition: str


def plain(text: str) -> str:
    return wenyan.strip_inline_tags(text).strip()


def clean_sanbu_lines(path: Path) -> list[str]:
    return [line for line in wenyan.clean_lines(path.read_text(encoding="utf-8")) if line.strip()]


def collapse_split_labels(lines: list[str]) -> list[str]:
    result: list[str] = []
    index = 0
    while index < len(lines):
        window = [plain(line) for line in lines[index : index + 4]]
        if window == ["常", "见", "义", "项"]:
            result.append("常见义项")
            index += 4
            continue
        if window == ["成", "语", "链", "接"]:
            result.append("成语链接")
            index += 4
            continue
        result.append(lines[index])
        index += 1
    return result


def add_vocab_example_icons(lines: list[str]) -> list[str]:
    result: list[str] = []
    in_common_items = False
    for line in lines:
        value = plain(line)
        if value == "常见义项":
            in_common_items = True
            result.append(line)
            continue
        if value == "成语链接" or re.match(r"^\d+[\.、]", value):
            in_common_items = False
        if (
            in_common_items
            and re.match(r"^[①②③④⑤⑥⑦⑧⑨⑩]", value)
            and not line.startswith(L_ICON_MARKER)
        ):
            result.append(L_ICON_MARKER + line)
            continue
        result.append(line)
    return result


def is_noise_volume_line(line: str) -> bool:
    value = plain(line)
    return value in SANBU_VOLUMES or value in SANBU_READING_LABELS


def next_nonempty(lines: list[str], index: int, count: int = 6) -> list[str]:
    result: list[str] = []
    for line in lines[index + 1 :]:
        if line.strip():
            result.append(line.strip())
        if len(result) >= count:
            break
    return result


def body_piece_start(lines: list[str], index: int, volume: str, reading_label: str) -> re.Match[str] | None:
    line = plain(lines[index])
    if not line.isdigit():
        return None
    num = int(line)
    if num <= 0 or num > 30:
        return None
    lookahead = [plain(item) for item in next_nonempty(lines, index, 6)]
    if len(lookahead) < 3:
        return None
    if re.fullmatch(r"\d+分钟", lookahead[0]):
        return None
    if lookahead[0] in SANBU_VOLUMES or lookahead[0] in SANBU_READING_LABELS:
        return None
    if volume == "进阶篇15篇":
        return re.match(r"^(\d+)$", line) if reading_label else None
    # Basic/promoted pieces have a source, word count, and time shortly after the title.
    if any(re.fullmatch(r"\d+分钟", item) for item in lookahead[2:6]):
        return re.match(r"^(\d+)$", line)
    if any("完成时间" in item for item in lookahead[2:6]):
        return re.match(r"^(\d+)$", line)
    return None


def piece_volume(base_volume: str, reading_label: str) -> str:
    if base_volume == "进阶篇15篇" and reading_label:
        return f"{base_volume} {reading_label}"
    return base_volume


def extract_body_sections(block: list[str]) -> tuple[str, dict[str, str], list[str], list[str]]:
    num_line = block[0]
    title = re.sub(r"[_＿]{2,}", "", plain(block[1])).strip() if len(block) > 1 else num_line
    content = [line for line in block[2:] if not is_noise_volume_line(line)]
    reading_index = 0
    for index, line in enumerate(content):
        value = plain(line)
        if value.startswith("阅读下面") or value.startswith("（20") or value.startswith("(20"):
            reading_index = index
            break
    vocab_index: int | None = None
    for index, line in enumerate(content):
        if plain(line) == "词语积累":
            vocab_index = index
            break

    meta_lines = content[:reading_index]
    main_end = vocab_index if vocab_index is not None else len(content)
    main_lines = content[reading_index:main_end]
    vocab_lines = content[vocab_index:] if vocab_index is not None else []
    vocab_lines = collapse_split_labels([line for line in vocab_lines if not is_noise_volume_line(line)])
    vocab_lines = add_vocab_example_icons(vocab_lines)

    meta_info = build_meta_info(meta_lines)
    return title, meta_info, main_lines, vocab_lines


def strip_meta_label(text: str, label: str) -> str:
    return re.sub(rf"^\s*{re.escape(label)}\s*[：:]\s*", "", text).strip()


def build_meta_info(meta_lines: list[str]) -> dict[str, str]:
    cleaned = [
        re.sub(r"[_＿]+", "", plain(line)).strip()
        for line in meta_lines
        if plain(line)
    ]
    cleaned = [line for line in cleaned if line]
    info = {"vocab": "", "source": "", "word_count": "", "time": ""}
    other: list[str] = []
    for item in cleaned:
        time_match = re.search(r"\d+\s*分钟", item)
        if time_match:
            info["time"] = time_match.group(0).replace(" ", "")
            continue
        if "字数" in item or re.fullmatch(r"\d+", item):
            count_match = re.search(r"\d+", item)
            info["word_count"] = count_match.group(0) if count_match else strip_meta_label(item, "字数")
            continue
        if "文章出处" in item:
            info["source"] = strip_meta_label(item, "文章出处")
            continue
        other.append(item)

    if other:
        info["vocab"] = strip_meta_label(other[0], "词语积累")
    if not info["source"] and len(other) >= 2:
        info["source"] = strip_meta_label(other[-1], "文章出处")
    return {key: value for key, value in info.items() if value}


def build_meta_summary(meta_lines: list[str]) -> str:
    info = build_meta_info(meta_lines)
    parts: list[str] = []
    if info.get("vocab"):
        parts.append(f"词语积累：{info['vocab']}")
    if info.get("source"):
        parts.append(f"文章出处：{info['source']}")
    if info.get("word_count"):
        parts.append(f"字数：{info['word_count']}")
    if info.get("time"):
        parts.append(f"完成时间：{info['time']}")
    return "　".join(parts)


def parse_body_pieces(md_path: Path) -> list[wenyan.Piece]:
    lines = clean_sanbu_lines(md_path)
    starts: list[tuple[int, str, int]] = []
    current_volume = ""
    current_reading_label = ""
    for index, line in enumerate(lines):
        value = plain(line)
        if value in SANBU_VOLUMES:
            current_volume = value
            current_reading_label = ""
            continue
        if value in SANBU_READING_LABELS:
            current_reading_label = value
            continue
        match = body_piece_start(lines, index, current_volume, current_reading_label)
        if match and current_volume:
            starts.append((index, piece_volume(current_volume, current_reading_label), int(match.group(1))))

    pieces: list[wenyan.Piece] = []
    for idx, (start, volume, num) in enumerate(starts):
        end = starts[idx + 1][0] if idx + 1 < len(starts) else len(lines)
        block = lines[start:end]
        title, meta_info, main_lines, vocab_lines = extract_body_sections(block)
        piece_lines = [title, MAIN_MARKER, *main_lines]
        if vocab_lines:
            piece_lines.extend([VOCAB_MARKER, *vocab_lines])
        pieces.append(
            wenyan.Piece(
                volume=volume,
                num=num,
                title=title,
                author=json.dumps(meta_info, ensure_ascii=False) if meta_info else "",
                lines=piece_lines,
            )
        )
    return pieces


def answer_piece_match(line: str) -> tuple[str, int, str] | None:
    value = plain(line)
    value = re.sub(r"[_＿]+", " ", value)
    match = re.match(r"^(基础|提升)\s*(\d+)\s*(.*)$", value)
    if match:
        volume = "基础篇30篇" if match.group(1) == "基础" else "提升篇20篇"
        return volume, int(match.group(2)), match.group(3).strip()
    match = re.match(r"^(单文本阅读|多文本阅读)\s*(\d+)\s*(.*)$", value)
    if match:
        return f"进阶篇15篇 {match.group(1)}", int(match.group(2)), match.group(3).strip()
    return None


def split_answer_sections(lines: list[str]) -> tuple[list[str], list[str], list[str]]:
    main: list[str] = []
    reference: list[str] = []
    vocab: list[str] = []
    in_vocab = False
    in_reference = False
    for line in lines:
        value = plain(line)
        if value in {"答案解析", *SANBU_VOLUMES}:
            continue
        if value == "【词语积累】":
            in_vocab = True
            in_reference = False
            continue
        if "参考译文" in line or "参考译文" in value:
            in_reference = True
            in_vocab = False
            reference.append(line)
            continue
        if in_reference:
            reference.append(line)
            continue
        if in_vocab:
            vocab.append(line)
        else:
            main.append(line)
    return main, reference, vocab


def parse_answer_pieces(md_path: Path, body_pieces: list[wenyan.Piece]) -> list[wenyan.Piece]:
    title_map = {(piece.volume, piece.num): piece.title for piece in body_pieces}
    lines = clean_sanbu_lines(md_path)
    starts: list[tuple[int, str, int, str]] = []
    for index, line in enumerate(lines):
        match = answer_piece_match(line)
        if match:
            volume, num, title = match
            starts.append((index, volume, num, title))

    pieces: list[wenyan.Piece] = []
    seen: set[tuple[str, int]] = set()
    for idx, (start, volume, num, title) in enumerate(starts):
        key = (volume, num)
        if key in seen:
            continue
        seen.add(key)
        end = starts[idx + 1][0] if idx + 1 < len(starts) else len(lines)
        block = lines[start + 1 : end]
        main, reference, vocab = split_answer_sections(block)
        resolved_title = re.sub(r"[_＿]{2,}", "", title).strip() or title_map.get(key, f"{num}")
        piece_lines = [resolved_title, MAIN_MARKER, *main]
        if reference:
            piece_lines.extend([REFERENCE_MARKER, *reference])
        if vocab:
            piece_lines.extend([VOCAB_MARKER, *vocab])
        pieces.append(wenyan.Piece(volume=volume, num=num, title=resolved_title, author="", lines=piece_lines))
    return pieces


def split_sanbu_piece_lines(piece: wenyan.Piece | None) -> dict[str, list[str]]:
    if piece is None:
        return {}
    sections: dict[str, list[str]] = {}
    current = ""
    buffer: list[str] = []
    for line in piece.lines[1:]:
        if line in {MAIN_MARKER, REFERENCE_MARKER, VOCAB_MARKER}:
            if current and buffer:
                sections[current] = buffer
            if line == MAIN_MARKER:
                current = "main"
            elif line == REFERENCE_MARKER:
                current = "reference"
            else:
                current = "vocab"
            buffer = []
            continue
        if current:
            buffer.append(line)
    if current and buffer:
        sections[current] = buffer
    return sections


def sanbu_split_analysis(text: str) -> tuple[str, list[str]]:
    matches = list(SANBU_ANALYSIS_LABEL_RE.finditer(text))
    if not matches:
        return text, []
    before = text[: matches[0].start()].strip()
    analyses: list[str] = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        analysis = text[match.start() : end].strip()
        if analysis and analysis != match.group(0):
            analyses.append(analysis)
    return before, analyses


def sanbu_meta_info(piece: wenyan.Piece) -> dict[str, str]:
    if not piece.author:
        return {}
    try:
        data = json.loads(piece.author)
    except json.JSONDecodeError:
        return {}
    if not isinstance(data, dict):
        return {}
    return {str(key): str(value) for key, value in data.items() if value}


def sanbu_icon_path(index: int) -> Path | None:
    for ext in (".png", ".jpg", ".jpeg"):
        candidate = SANBU_STATIC_DIR / f"{index}{ext}"
        if candidate.exists():
            return candidate
    return None


def sanbu_icon_package_path(index: int) -> str:
    return f"ppt/media/sanbu_meta_icon_{index}.png"


def append_sanbu_icon(
    root,
    rels_root,
    media_assets: dict[str, bytes],
    *,
    index: int,
    x: float,
    y: float,
    size: float = 0.24,
) -> None:
    icon_path = sanbu_icon_path(index)
    if icon_path is None:
        return
    package_path = sanbu_icon_package_path(index)
    if package_path not in media_assets:
        media_assets[package_path] = wenyan.image_to_png_bytes(icon_path)
    rel_id = wenyan.add_image_relationship(rels_root, wenyan.posixpath.relpath(package_path, "ppt/slides"))
    wenyan.append_picture(
        root,
        rel_id=rel_id,
        x=x,
        y=y,
        w=size,
        h=size,
        name=f"三步进阶信息图标{index}",
    )


def sanbu_text_width(text: str, *, min_w: float, max_w: float, scale: float = 0.22) -> float:
    return max(min_w, min(max_w, wenyan.text_units(text) * scale))


def append_sanbu_meta_text(root, *, x: float, y: float, w: float, text: str, role: str) -> None:
    h = 0.54 if role in {"sanbu_meta_top", "sanbu_meta_info"} else 0.44
    wenyan.append_textbox(
        root,
        x=x,
        y=y,
        w=w,
        h=h,
        lines=[wenyan.StyledLine(text, role)],
        name=f"三步进阶元信息 {text[:8]}",
    )


def set_shape_box(sp, *, x: float, y: float, w: float, h: float) -> None:
    xfrm = sp.find(f".//{{{wenyan.A_NS}}}xfrm")
    if xfrm is None:
        return
    off = xfrm.find(wenyan.qn(wenyan.A_NS, "off"))
    ext = xfrm.find(wenyan.qn(wenyan.A_NS, "ext"))
    if off is not None:
        off.attrib["x"] = str(wenyan.inch(x))
        off.attrib["y"] = str(wenyan.inch(y))
    if ext is not None:
        ext.attrib["cx"] = str(wenyan.inch(w))
        ext.attrib["cy"] = str(wenyan.inch(h))


def set_text_shape_size(sp, size: float) -> None:
    size_value = str(int(size * 100))
    for element in sp.findall(f".//{{{wenyan.A_NS}}}rPr") + sp.findall(f".//{{{wenyan.A_NS}}}endParaRPr"):
        element.attrib["sz"] = size_value


def append_sanbu_meta_header(root, rels_root, media_assets: dict[str, bytes], meta: dict[str, str]) -> None:
    icon_size = 0.36
    gap = 0.10
    vocab_text = f"词语积累：{meta.get('vocab', '')}".strip("：")
    if meta.get("vocab"):
        text_w = sanbu_text_width(vocab_text, min_w=3.15, max_w=7.60, scale=0.30)
        group_w = icon_size + gap + text_w
        x = (wenyan.SLIDE_W_IN - group_w) / 2
        append_sanbu_icon(root, rels_root, media_assets, index=1, x=x, y=1.36, size=icon_size)
        append_sanbu_meta_text(root, x=x + icon_size + gap, y=1.23, w=text_w, text=vocab_text, role="sanbu_meta_top")

    second_icon_y = 1.92
    second_text_y = 1.79
    source_text = f"文章出处：{meta.get('source', '')}" if meta.get("source") else ""
    word_text = f"字数：{meta.get('word_count', '')}" if meta.get("word_count") else ""
    time_text = f"完成时间：{meta.get('time', '')}" if meta.get("time") else ""
    groups = [
        (2, source_text, 1.08, 6.35),
        (3, word_text, 7.25, 1.50),
        (4, time_text, 9.45, 3.05),
    ]
    for icon_index, text, x, width in groups:
        if not text:
            continue
        append_sanbu_icon(root, rels_root, media_assets, index=icon_index, x=x, y=second_icon_y, size=icon_size)
        append_sanbu_meta_text(
            root,
            x=x + icon_size + gap,
            y=second_text_y,
            w=width,
            text=text,
            role="sanbu_meta_info",
        )


def append_sanbu_answer_item_text(
    item: wenyan.AnswerItem,
    text: str,
    extract_ordered_answers: Callable[[list[str]], list[str]],
) -> None:
    answer_part, analyses = sanbu_split_analysis(text.strip())
    if answer_part:
        item.answers.extend(extract_ordered_answers([answer_part]))
    item.analyses.extend(analyses)


def split_circled_segments(text: str) -> list[str]:
    pattern = re.compile(r"[①②③④⑤⑥⑦⑧⑨⑩]")
    matches = list(pattern.finditer(text))
    if not matches:
        return [text.strip()] if text.strip() else []
    segments: list[str] = []
    for index, match in enumerate(matches):
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        segment = text[start:end].strip()
        if segment:
            segments.append(segment)
    return segments


def split_meaning_and_definition(text: str) -> list[str]:
    text = text.strip()
    if not text:
        return []
    parts = re.split(r"(?<=。)", text, maxsplit=1)
    if len(parts) == 1:
        pieces = re.split(r"[；;]", text, maxsplit=1)
        return [piece.strip() for piece in pieces if piece.strip()]
    return [part.strip() for part in parts if part.strip()]


def is_sanbu_glossary_reading(lines: list[str], answer_lines: list[str] | None) -> bool:
    if not answer_lines:
        return False
    return any(SANBU_GLOSSARY_PROMPT in plain(line) for line in lines) and any(
        plain(line) == "【重点字词】" for line in answer_lines
    )


def sanbu_glossary_answer_payload(answer_lines: list[str] | None) -> str:
    if not answer_lines:
        return ""
    collecting = False
    payload: list[str] = []
    for raw_line in answer_lines:
        value = plain(raw_line)
        if value == "【重点字词】":
            collecting = True
            continue
        if not collecting:
            continue
        if not value:
            continue
        if wenyan.answer_item_match(value):
            break
        if value in SANBU_VOLUMES or value in SANBU_READING_LABELS:
            continue
        payload.append(value)
    return " ".join(payload)


def parse_sanbu_glossary_entries(answer_lines: list[str] | None) -> list[SanbuGlossaryEntry]:
    payload = sanbu_glossary_answer_payload(answer_lines)
    if not payload:
        return []
    entry_re = re.compile(rf"([{re.escape(SANBU_CIRCLED_NUMBERS)}]?)([^：:，。；;\s]{{1,10}})[：:]")
    matches = list(entry_re.finditer(payload))
    entries: list[SanbuGlossaryEntry] = []
    for index, match in enumerate(matches):
        term = match.group(2).strip()
        if not term:
            continue
        end = matches[index + 1].start() if index + 1 < len(matches) else len(payload)
        definition = payload[match.end() : end].strip(" \t\r\n；;")
        if not definition:
            continue
        entries.append(SanbuGlossaryEntry(label=match.group(1), term=term, definition=definition))
    return entries


def remove_sanbu_glossary_blank_tokens(text: str) -> str:
    text = re.sub(r'<span\s+class="blank"[^>]*>\s*</span>', "", text)
    text = re.sub(r'<span\s+class="uline"[^>]*>[\s_＿　]+</span>', "", text)
    text = re.sub(r"[_＿]{2,}", "", text)
    return text


def sanbu_glossary_answer_html(entry: SanbuGlossaryEntry) -> str:
    answer = f"{entry.term}：{entry.definition}"
    return f"{entry.label}（{wenyan.answer_span(answer)}）"


def apply_sanbu_glossary_entries_to_text(text: str, entries: list[SanbuGlossaryEntry]) -> str:
    text = remove_sanbu_glossary_blank_tokens(text)
    cursor = 0
    for entry in entries:
        term = entry.term
        if not term:
            continue
        answer_html = sanbu_glossary_answer_html(entry)
        term_pos = text.find(term, cursor)
        if term_pos >= 0:
            term_end = term_pos + len(term)
            label_pos = -1
            if entry.label:
                label_pos = text.find(entry.label, term_end, min(len(text), term_end + 160))
            blue_term = f'<span class="blue-text">{html.escape(term)}</span>'
            if label_pos >= 0:
                between = remove_sanbu_glossary_blank_tokens(text[term_end:label_pos])
                text = text[:term_pos] + blue_term + answer_html + between + text[label_pos + len(entry.label) :]
                cursor = term_pos + len(blue_term) + len(answer_html)
            else:
                text = text[:term_pos] + blue_term + answer_html + text[term_end:]
                cursor = term_pos + len(blue_term) + len(answer_html)
            continue

        # Some rare glyphs are stored as tif pictures in the md (for example
        # juan.tif + juan.tif for "惓惓"). In that case the word cannot be
        # searched as text, so anchor the answer to its original circled marker.
        if entry.label:
            label_pos = text.find(entry.label, cursor)
            if label_pos >= 0:
                text = text[:label_pos] + answer_html + text[label_pos + len(entry.label) :]
                cursor = label_pos + len(answer_html)
    return text


def sanbu_html_visible_text(text: str) -> str:
    text = re.sub(r"<img\b[^>]*>", "\u3000", text, flags=re.I)
    return html.unescape(re.sub(r"<[^>]+>", "", text))


def sanbu_html_visible_units(text: str) -> float:
    return wenyan.text_units(sanbu_html_visible_text(text))


def iter_sanbu_html_atoms(text: str):
    index = 0
    while index < len(text):
        if text[index] == "<":
            end = text.find(">", index)
            if end >= 0:
                tag = text[index : end + 1]
                yield tag, 1.0 if re.match(r"<img\b", tag, flags=re.I) else 0.0
                index = end + 1
                continue
        yield text[index], 1.0 if ord(text[index]) > 127 else 0.55
        index += 1


def split_sanbu_glossary_html_fragments(text: str, max_units: float = 82.0) -> list[str]:
    fragments: list[str] = []
    current = ""
    units = 0.0
    line_units = 0.0
    last_safe = -1
    safe_chars = "。！？；;）”"
    answer_group_re = re.compile(
        rf"(<span class=\"blue-text\">.*?</span>[{re.escape(SANBU_CIRCLED_NUMBERS)}]?（<span class=\"answer-fill\">.*?</span>）|"
        rf"(?:<img\b[^>]*>\s*)+[{re.escape(SANBU_CIRCLED_NUMBERS)}]（<span class=\"answer-fill\">.*?</span>）|"
        rf"[{re.escape(SANBU_CIRCLED_NUMBERS)}]（<span class=\"answer-fill\">.*?</span>）)",
        re.S,
    )

    def emit(fragment: str) -> None:
        fragment = fragment.strip()
        if fragment:
            fragments.append(fragment)

    def reset_with(fragment: str) -> None:
        nonlocal current, units, line_units, last_safe
        current = fragment.lstrip()
        units = sanbu_html_visible_units(current)
        line_units = units % 32.0
        last_safe = -1

    def flush_current() -> None:
        nonlocal current, units, line_units, last_safe
        emit(current)
        current = ""
        units = 0.0
        line_units = 0.0
        last_safe = -1

    def append_atom(atom: str, atom_units: float) -> None:
        nonlocal current, units, line_units, last_safe
        current += atom
        units += atom_units
        line_units = (line_units + atom_units) % 32.0
        if atom in safe_chars:
            last_safe = len(current)
        if units >= max_units:
            if 0 < last_safe < len(current):
                emit(current[:last_safe])
                reset_with(current[last_safe:])
            else:
                flush_current()

    index = 0
    for match in answer_group_re.finditer(text):
        normal = text[index : match.start()]
        for atom, atom_units in iter_sanbu_html_atoms(normal):
            append_atom(atom, atom_units)

        group = match.group(0)
        group_units = sanbu_html_visible_units(group)
        if current.strip() and line_units > 0 and (
            line_units + group_units > 24.0
            or (group_units >= 8.0 and line_units > 8.0)
            or (group_units >= 4.5 and units > 8.0)
        ):
            flush_current()
        current += group
        units += group_units
        line_units = (line_units + group_units) % 32.0
        last_safe = len(current)
        if units >= max_units:
            flush_current()
        index = match.end()

    for atom, atom_units in iter_sanbu_html_atoms(text[index:]):
        append_atom(atom, atom_units)
    flush_current()

    separated: list[str] = []
    pending = [fragment for fragment in fragments if fragment.strip()]
    while pending:
        fragment = pending.pop(0)
        matches = list(answer_group_re.finditer(fragment))
        if len(matches) <= 1:
            separated.append(fragment)
            continue
        split_at = matches[1].start()
        head = fragment[:split_at].strip()
        tail = fragment[split_at:].strip()
        if head:
            separated.append(head)
        if tail:
            pending.insert(0, tail)
    return separated or [text]


def build_sanbu_glossary_preamble_lines(
    leading: list[str],
    answer_lines: list[str] | None,
) -> list[str]:
    entries = parse_sanbu_glossary_entries(answer_lines)
    prompt_lines: list[str] = []
    passage_parts: list[str] = []
    tail_lines: list[str] = []
    in_tail = False
    for raw_line in leading:
        value = plain(raw_line)
        if not value:
            continue
        if SANBU_GLOSSARY_PROMPT in value:
            prompt_lines.append(raw_line)
            continue
        if value.startswith("【注】") or wenyan.is_source_note_plain(value):
            in_tail = True
        if in_tail:
            tail_lines.append(remove_sanbu_glossary_blank_tokens(raw_line))
        else:
            passage_parts.append(raw_line)

    result = list(prompt_lines)
    if passage_parts:
        passage_text = "".join(remove_sanbu_glossary_blank_tokens(line) for line in passage_parts)
        filled_passage = apply_sanbu_glossary_entries_to_text(passage_text, entries)
        result.extend(split_sanbu_glossary_html_fragments(filled_passage))
    result.extend(tail_lines)
    return [line for line in result if plain(line)]


def vocab_answer_values(answer_lines: list[str] | None) -> list[str]:
    values: list[str] = []
    if not answer_lines:
        return values
    for raw_line in answer_lines:
        line = plain(raw_line)
        if not line:
            continue
        if line.startswith("成语链接"):
            payload = re.sub(r"^成语链接\s*[：:]\s*", "", line).strip()
            for segment in split_circled_segments(payload):
                values.extend(split_meaning_and_definition(segment))
            continue
        match = re.match(r"^\d+[\.、]\s*(.*)$", line)
        if match:
            values.extend(split_circled_segments(match.group(1)))
    return values


def normalize_sanbu_fillable_blanks(line: str) -> str:
    line = re.sub(r'<span\s+class="blank"[^>]*>\s*</span>', "____", line)
    line = re.sub(r'<span\s+class="uline"[^>]*>[\s_＿]+</span>', "____", line)
    return line


def strip_vocab_answer_padding(line: str) -> str:
    return re.sub(
        r'(<span\s+class="answer-fill(?:\s+uline)?">)\u3000(.*?)\u3000(</span>)',
        r"\1\2\3",
        line,
    )


def add_blue_reading_underline_classes(line: str) -> str:
    def repl(match: re.Match[str]) -> str:
        classes = match.group(1).split()
        if ("uline" in classes or "wave" in classes) and "blue-underline" not in classes:
            classes.append("blue-underline")
        return f'class="{" ".join(classes)}"'

    return re.sub(r'class="([^"]*\b(?:uline|wave)\b[^"]*)"', repl, line)


def mark_reading_passage_underlines_blue(lines: list[str]) -> list[str]:
    result: list[str] = []
    in_reading_passage = False
    for line in lines:
        value = plain(line)
        if "阅读下面的文言文" in value:
            in_reading_passage = True
        elif in_reading_passage and re.match(r"^\d+[\.．、]\s*", value):
            in_reading_passage = False

        if in_reading_passage:
            result.append(add_blue_reading_underline_classes(line))
        else:
            result.append(line)
    return result


def sanbu_underlined_spaces(units: int) -> str:
    return f'<span class="uline">{"　" * max(1, units)}</span>'


def sanbu_inline_answer_line(answer: str, *, total_units: int) -> str:
    answer = answer.strip()
    if not answer:
        return sanbu_underlined_spaces(total_units)
    answer_units = max(1, int(wenyan.text_units(answer) + 0.999))
    padding_units = max(2, total_units - answer_units)
    left_units = max(1, padding_units // 2)
    right_units = max(1, padding_units - left_units)
    return (
        f"{sanbu_underlined_spaces(left_units)}"
        f"{wenyan.answer_span(answer, underline=True)}"
        f"{sanbu_underlined_spaces(right_units)}"
    )


def sanbu_centered_answer_marker(answer: str, *, total_units: int) -> str:
    answer = answer.strip()
    if not answer:
        return ""
    answer_units = max(1, int(wenyan.text_units(answer) + 0.999))
    left_units = max(0, (total_units - answer_units) // 2)
    return f"{'　' * left_units}{wenyan.answer_span(answer)}"


ANSWER_FILL_RE = re.compile(r'<span\s+class="answer-fill(?:\s+uline)?">(.*?)</span>')


def extract_answer_fill(text: str) -> tuple[str, str, str] | None:
    match = ANSWER_FILL_RE.search(text)
    if not match:
        return None
    answer = html.unescape(wenyan.strip_inline_tags(match.group(1))).strip()
    return text[: match.start()], answer, text[match.end() :]


def split_idiom_term_and_label(prefix_html: str) -> tuple[str, str]:
    value = plain(prefix_html).rstrip()
    if len(value) >= 2 and value[-1] in {"：", ":"}:
        label_char = value[-2]
        colon = value[-1]
        token = f"{label_char}{colon}"
        if token in prefix_html:
            term_html, _tail = prefix_html.rsplit(token, 1)
            return term_html.strip(), f"{label_char}："
    return prefix_html.strip(), ""


def format_sanbu_idiom_item(item_line: str, definition_line: str | None) -> list[wenyan.StyledLine]:
    parsed_item = extract_answer_fill(item_line)
    if parsed_item is None:
        return [wenyan.StyledLine(item_line, "body")]

    item_prefix, item_answer, item_suffix = parsed_item
    term_html, label = split_idiom_term_and_label(item_prefix)
    term_units = wenyan.text_units(plain(term_html))
    gap_units = max(2, int(16 - term_units))
    item_text = (
        f"{term_html}{'　' * gap_units}{label}"
        f"{sanbu_inline_answer_line(item_answer + item_suffix.strip(), total_units=16)}"
    )
    result = [wenyan.StyledLine(item_text, "body")]

    if definition_line:
        parsed_definition = extract_answer_fill(definition_line)
        if parsed_definition is not None:
            definition_prefix, definition_answer, definition_suffix = parsed_definition
            definition_text = (
                f"{definition_prefix}"
                f"{sanbu_inline_answer_line(definition_answer + definition_suffix.strip(), total_units=30)}"
            )
            result.append(wenyan.StyledLine(definition_text, "body"))
        else:
            result.append(wenyan.StyledLine(definition_line, "body"))
    return result


def format_sanbu_idiom_links(styled: list[wenyan.StyledLine]) -> list[wenyan.StyledLine]:
    result: list[wenyan.StyledLine] = []
    index = 0
    while index < len(styled):
        line = styled[index]
        if plain(line.text) != "成语链接":
            result.append(line)
            index += 1
            continue

        result.append(wenyan.StyledLine("成语链接", "heading"))
        index += 1
        while index < len(styled):
            current = styled[index]
            current_plain = plain(current.text)
            if current_plain == "成语链接" or re.match(r"^\d+[\.．、]\s*", current_plain):
                break

            next_line = styled[index + 1] if index + 1 < len(styled) else None
            next_plain = plain(next_line.text) if next_line is not None else ""
            if next_line is not None and next_plain.startswith("成语释义"):
                result.extend(format_sanbu_idiom_item(current.text, next_line.text))
                index += 2
                continue

            result.append(current)
            index += 1
    return result


def build_sanbu_vocab_styled_section(
    body_lines: list[str],
    answer_lines: list[str] | None,
) -> tuple[list[wenyan.StyledLine], bool]:
    answers = vocab_answer_values(answer_lines)
    body_lines = [normalize_sanbu_fillable_blanks(line) for line in body_lines]
    filled_lines, matched = wenyan.fill_lines_with_answer_values(body_lines, answers)
    filled_lines = [strip_vocab_answer_padding(line) for line in filled_lines]
    cleaned_lines: list[str] = []
    icon_line_indexes: set[int] = set()
    pending_icon = False
    for line in filled_lines:
        if line.startswith(L_ICON_MARKER):
            pending_icon = True
            cleaned_lines.append(line[len(L_ICON_MARKER) :])
        elif pending_icon and plain(line).startswith("义项"):
            icon_line_indexes.add(len(cleaned_lines))
            cleaned_lines.append(
                '<img class="fbd-tif-char" src="L.tif" alt="L.tif" data-fbd-tif="L.tif" />'
                + line
            )
            pending_icon = False
        else:
            cleaned_lines.append(line)
    styled = wenyan.styled_body_lines(cleaned_lines, cleaned_lines[0] if cleaned_lines else "")
    for index, line in enumerate(styled):
        if index in icon_line_indexes:
            meta = dict(line.meta or {})
            meta["l_icon"] = True
            line.meta = meta
    styled = format_sanbu_idiom_links(styled)
    return styled, matched


def append_vocab_l_icons(
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
    tif_path = wenyan.resolve_tif_path("L.tif", tif_search_dirs)
    if tif_path is None:
        return
    cursor = 0.0
    for line in lines:
        role = line.role or "body"
        line_h = wenyan.line_height_for_role(role)
        plain_text = plain(line.text)
        visual_lines = max(1, int((wenyan.text_units(plain_text) + 23.0) // 24.0))
        if line.meta and line.meta.get("l_icon"):
            wenyan.append_tif_picture(
                root,
                rels_root,
                media_assets,
                tif_path=tif_path,
                x=x + 0.69,
                y=y + cursor + 0.03,
                h=0.34,
                name="词语积累例图标",
            )
        cursor += visual_lines * line_h


def append_sanbu_solid_line(
    root,
    *,
    x: float,
    y: float,
    w: float,
    name: str = "三步进阶长答案横线",
) -> int:
    shape_id = wenyan.next_shape_id(root)
    cxn = wenyan.ET.Element(wenyan.qn(wenyan.P_NS, "cxnSp"))
    nv_pr = wenyan.ET.SubElement(cxn, wenyan.qn(wenyan.P_NS, "nvCxnSpPr"))
    wenyan.ET.SubElement(nv_pr, wenyan.qn(wenyan.P_NS, "cNvPr"), {"id": str(shape_id), "name": f"{name} {shape_id}"})
    wenyan.ET.SubElement(nv_pr, wenyan.qn(wenyan.P_NS, "cNvCxnSpPr"))
    wenyan.ET.SubElement(nv_pr, wenyan.qn(wenyan.P_NS, "nvPr"))

    sp_pr = wenyan.ET.SubElement(cxn, wenyan.qn(wenyan.P_NS, "spPr"))
    xfrm = wenyan.ET.SubElement(sp_pr, wenyan.qn(wenyan.A_NS, "xfrm"))
    wenyan.ET.SubElement(xfrm, wenyan.qn(wenyan.A_NS, "off"), {"x": str(wenyan.inch(x)), "y": str(wenyan.inch(y))})
    wenyan.ET.SubElement(xfrm, wenyan.qn(wenyan.A_NS, "ext"), {"cx": str(wenyan.inch(w)), "cy": "0"})
    geom = wenyan.ET.SubElement(sp_pr, wenyan.qn(wenyan.A_NS, "prstGeom"), {"prst": "line"})
    wenyan.ET.SubElement(geom, wenyan.qn(wenyan.A_NS, "avLst"))
    ln = wenyan.ET.SubElement(sp_pr, wenyan.qn(wenyan.A_NS, "ln"), {"w": "9525"})
    solid = wenyan.ET.SubElement(ln, wenyan.qn(wenyan.A_NS, "solidFill"))
    wenyan.ET.SubElement(solid, wenyan.qn(wenyan.A_NS, "srgbClr"), {"val": "000000"})
    wenyan.append_to_sp_tree(root, cxn)
    return shape_id


def sanbu_visual_line_count(line: wenyan.StyledLine) -> int:
    if (line.role or "") == "long_answer":
        return 1
    return max(1, wenyan.estimate_line_count(line.text))


def append_sanbu_long_answer_underlines(
    root,
    *,
    x: float,
    y: float,
    w: float,
    lines: list[wenyan.StyledLine],
) -> None:
    cursor = 0.0
    for line in lines:
        role = line.role or "body"
        line_h = wenyan.line_height_for_role(role)
        if role == "long_answer":
            append_sanbu_solid_line(
                root,
                x=x + 0.70,
                y=y + cursor + min(0.60, line_h * 1.50),
                w=max(1.0, w - 1.05),
            )
        cursor += sanbu_visual_line_count(line) * line_h


def render_sanbu_text_content(
    root,
    rels_root,
    media_assets: dict[str, bytes],
    *,
    piece: wenyan.Piece,
    lines: list[wenyan.StyledLine],
    tif_search_dirs: list[Path],
    x: float,
    y: float,
    w: float,
    h: float,
    name_suffix: str = "content",
) -> None:
    wenyan.append_textbox(
        root,
        x=x,
        y=y,
        w=w,
        h=h,
        lines=lines,
        name=f"{piece.title} {name_suffix}",
        answer_mode="base",
    )
    if any((line.role or "") == "long_answer" for line in lines):
        append_sanbu_long_answer_underlines(root, x=x, y=y, w=w, lines=lines)
    wenyan.append_tif_overlays(
        root,
        rels_root,
        media_assets,
        x=x,
        y=y,
        w=w,
        lines=lines,
        tif_search_dirs=tif_search_dirs,
    )
    animations = wenyan.append_answer_overlays(
        root,
        x=x,
        y=y,
        w=w,
        h=h,
        lines=lines,
        name=f"{piece.title} {name_suffix}",
    )
    wenyan.set_slide_animations(root, animations)


def is_sanbu_numbered_prompt(line: wenyan.StyledLine) -> bool:
    if (line.role or "body") not in {"body", "reading", "extended_reading"}:
        return False
    text = plain(line.text)
    return bool(re.match(r"^\d+[\.．、]\s*", text))


def repair_orphan_sanbu_question_chunks(
    chunks: list[list[wenyan.StyledLine]],
) -> list[list[wenyan.StyledLine]]:
    repaired: list[list[wenyan.StyledLine]] = []
    index = 0
    while index < len(chunks):
        current = list(chunks[index])
        if (
            index + 1 < len(chunks)
            and current
            and len(current) >= 2
            and is_sanbu_numbered_prompt(current[-1])
            and (chunks[index + 1][0].role or "body") in {"body", "reading", "extended_reading", "long_answer"}
        ):
            carry = current.pop()
            if current:
                repaired.append(current)
            chunks[index + 1] = [carry] + chunks[index + 1]
            index += 1
            continue
        repaired.append(current)
        index += 1
    return [chunk for chunk in repaired if chunk]


def split_mixed_segmentation_chunks(
    chunks: list[list[wenyan.StyledLine]],
) -> list[list[wenyan.StyledLine]]:
    result: list[list[wenyan.StyledLine]] = []
    segmentation_roles = {"segmentation_question", "segmentation_text", "segmentation_analysis"}
    for chunk in chunks:
        if not any((line.role or "") == "segmentation_text" for line in chunk):
            result.append(chunk)
            continue
        current: list[wenyan.StyledLine] = []
        index = 0
        while index < len(chunk):
            line = chunk[index]
            if (line.role or "") == "segmentation_question":
                if current:
                    result.append(current)
                    current = []
                seg_block: list[wenyan.StyledLine] = []
                while index < len(chunk) and (chunk[index].role or "") in segmentation_roles:
                    seg_block.append(chunk[index])
                    index += 1
                if seg_block:
                    result.append(seg_block)
                continue
            current.append(line)
            index += 1
        if current:
            result.append(current)
    return result


def sanbu_chunk_cost(chunk: list[wenyan.StyledLine]) -> int:
    return sum(wenyan.line_page_cost(line) for line in chunk)


def merge_split_long_answer_chunks(
    chunks: list[list[wenyan.StyledLine]],
    *,
    max_units: int = 12,
) -> list[list[wenyan.StyledLine]]:
    result: list[list[wenyan.StyledLine]] = []
    for chunk in chunks:
        current = list(chunk)
        previous_can_take_answer = bool(
            result
            and result[-1]
            and (
                (result[-1][-1].role or "") == "long_answer"
                or (
                    any(is_sanbu_numbered_prompt(line) for line in result[-1])
                    and not any((line.role or "") in {"answer", "long_answer", "answer_spacer", "analysis"} for line in result[-1])
                )
            )
        )
        if (
            previous_can_take_answer
            and current
            and (current[0].role or "") == "long_answer"
        ):
            carry: list[wenyan.StyledLine] = []
            while current and (current[0].role or "") == "long_answer":
                carry.append(current.pop(0))
            optional: list[wenyan.StyledLine] = []
            while current and (current[0].role or "") in {"answer_spacer", "analysis"}:
                optional.append(current.pop(0))
            if sanbu_chunk_cost(result[-1] + carry + optional) <= max_units + 2:
                result[-1].extend(carry + optional)
            else:
                result[-1].extend(carry)
                current = optional + current
            if current:
                result.append(current)
            continue
        result.append(current)
    return [chunk for chunk in result if chunk]


def split_numbered_prompt_after_answer_chunks(
    chunks: list[list[wenyan.StyledLine]],
) -> list[list[wenyan.StyledLine]]:
    answer_roles = {"answer", "long_answer", "answer_spacer", "analysis", "segmentation_analysis"}
    result: list[list[wenyan.StyledLine]] = []
    for chunk in chunks:
        current: list[wenyan.StyledLine] = []
        seen_answer_content = False
        for line in chunk:
            if current and seen_answer_content and is_sanbu_numbered_prompt(line):
                result.append(current)
                current = [line]
                seen_answer_content = False
                continue
            current.append(line)
            role = line.role or "body"
            if role in answer_roles or "answer-fill" in line.text:
                seen_answer_content = True
        if current:
            result.append(current)
    return [chunk for chunk in result if chunk]


def add_slash_answer_labels(styled: list[wenyan.StyledLine]) -> list[wenyan.StyledLine]:
    result: list[wenyan.StyledLine] = []
    awaiting_slash_answer = False
    for line in styled:
        text = plain(line.text)
        if "请用" in text and "断句" in text:
            awaiting_slash_answer = True
            result.append(line)
            continue
        if awaiting_slash_answer and line.role == "answer" and "/" in text and not text.startswith("【答案】"):
            result.append(wenyan.StyledLine(f"【答案】{line.text}", "analysis", line.meta))
            awaiting_slash_answer = False
            continue
        if awaiting_slash_answer and line.role == "answer" and not text.startswith("【答案】"):
            result.append(wenyan.StyledLine(f"【答案】{line.text}", "analysis", line.meta))
            awaiting_slash_answer = False
            continue
        if awaiting_slash_answer and line.role in {"analysis", "heading", "subheading", "long_answer"}:
            awaiting_slash_answer = False
        result.append(line)
    return result


def chunk_sanbu_meta_main_lines(lines: list[wenyan.StyledLine]) -> list[list[wenyan.StyledLine]]:
    first_pass = wenyan.chunk_styled_lines(lines, max_units=6)
    if len(first_pass) <= 1:
        return first_pass
    rest_lines = [line for chunk in first_pass[1:] for line in chunk]
    rest_chunks = wenyan.chunk_styled_lines(rest_lines, max_units=12)
    rest_chunks = split_mixed_segmentation_chunks(repair_orphan_sanbu_question_chunks(rest_chunks))
    rest_chunks = split_numbered_prompt_after_answer_chunks(rest_chunks)
    rest_chunks = merge_split_long_answer_chunks(rest_chunks)
    rest_chunks = split_numbered_prompt_after_answer_chunks(rest_chunks)
    return [first_pass[0]] + rest_chunks


def split_sanbu_long_answer(answer: str, max_units: float = 34.0) -> list[str]:
    pieces: list[str] = []
    current = ""
    units = 0.0
    last_break = -1
    break_chars = "，。；、！？,;；"
    for ch in answer:
        weight = 1.0 if ord(ch) > 127 else 0.55
        if current and units + weight > max_units:
            if last_break >= 0 and wenyan.text_units(current[:last_break]) >= max_units * 0.55:
                pieces.append(current[:last_break].strip())
                current = current[last_break:].lstrip()
            else:
                pieces.append(current.strip())
                current = ""
            units = wenyan.text_units(current)
            last_break = -1
        current += ch
        units += weight
        if ch in break_chars:
            last_break = len(current)
    if current.strip():
        pieces.append(current.strip())
    return pieces[:4]


def sanbu_fixed_answer_line(answer: str = "") -> str:
    return sanbu_centered_answer_marker(answer, total_units=34)


def sanbu_fixed_long_answer_lines(answer: str) -> list[wenyan.StyledLine]:
    pieces = split_sanbu_long_answer(answer)
    if not pieces:
        pieces = [""]
    line_count = len(pieces)
    return [
        wenyan.StyledLine(sanbu_fixed_answer_line(pieces[index] if index < len(pieces) else ""), "long_answer")
        for index in range(line_count)
    ]


def is_sanbu_translation_group(group_lines: list[str]) -> bool:
    group_text = "\n".join(plain(line) for line in group_lines)
    return "翻译成现代汉语" in group_text and bool(re.search(r"[（(]\s*1\s*[）)]", group_text))


def strip_sanbu_subanswer_prefix(answer: str) -> str:
    return re.sub(r"^[（(]\s*\d+\s*[）)]\s*", "", answer).strip()


def split_sanbu_translation_answers(item: wenyan.AnswerItem) -> list[str]:
    answers: list[str] = []
    for answer in item.answers:
        value = plain(answer)
        if SANBU_ANALYSIS_LABEL_RE.match(value):
            continue
        answers.append(strip_sanbu_subanswer_prefix(answer))
    return answers


def format_sanbu_translation_group(
    group_lines: list[str],
    item: wenyan.AnswerItem,
    section_title: str,
) -> list[wenyan.StyledLine]:
    result: list[wenyan.StyledLine] = []
    sub_answers = split_sanbu_translation_answers(item)
    answer_index = 0
    pending_question = False
    for index, line in enumerate(group_lines):
        clean_line = wenyan.strip_fillable_blanks_from_line(normalize_sanbu_fillable_blanks(line))
        value = plain(clean_line)
        if not value:
            continue
        if index == 0:
            result.extend(wenyan.styled_body_lines([clean_line], section_title))
            continue
        if re.match(r"^[（(]\s*\d+\s*[）)]", value):
            result.extend(wenyan.styled_body_lines([clean_line], section_title))
            if answer_index < len(sub_answers):
                result.extend(sanbu_fixed_long_answer_lines(sub_answers[answer_index]))
                answer_index += 1
            pending_question = True
            continue
        if re.fullmatch(r"[_＿]+", value):
            continue
        result.extend(wenyan.styled_body_lines([clean_line], section_title))
        pending_question = False

    if item.analyses:
        result.append(wenyan.StyledLine("", "answer_spacer"))
        for analysis in item.analyses:
            result.append(wenyan.StyledLine(analysis, "analysis"))
    return result


def is_sentence_pattern_group(group_lines: list[str]) -> bool:
    plain_group = "\n".join(plain(line) for line in group_lines)
    return "判断文中画波浪线的句子属于哪种文言句式" in plain_group and "句式类型" in plain_group and "句意" in plain_group


def normalize_sentence_pattern_line(line: str, item: wenyan.AnswerItem) -> str:
    answers = list(item.answers)
    sentence_type = answers[0] if len(answers) >= 1 else ""
    meaning = answers[1] if len(answers) >= 2 else ""
    same_sentence = answers[2] if len(answers) >= 3 else ""

    value = re.sub(r'<span\s+class="blank"[^>]*>\s*</span>', "", line)
    stripped = plain(value)
    stripped_no_blanks = re.sub(r"[_＿]+", "", stripped).strip()
    if "句式类型" in stripped_no_blanks and "句意" in stripped_no_blanks:
        sentence = re.split(r"句式类型", stripped_no_blanks, maxsplit=1)[0].strip()
        return (
            f"{sentence}　　句式类型：（{wenyan.answer_span(sentence_type)}）　"
            f"句意：{wenyan.answer_span(meaning, underline=True)}"
        )
    if stripped_no_blanks.startswith("请回忆学过的课文"):
        prefix = re.sub(r"[:：]\s*$", "：", stripped_no_blanks)
        return f"{prefix}{wenyan.answer_span(same_sentence, underline=True)}"
    return value


def format_sentence_pattern_group(
    group_lines: list[str],
    item: wenyan.AnswerItem,
    section_title: str,
) -> list[wenyan.StyledLine]:
    result: list[wenyan.StyledLine] = []
    answers = list(item.answers)
    sentence_type = answers[0] if len(answers) >= 1 else ""
    meaning = answers[1] if len(answers) >= 2 else ""
    same_sentence = answers[2] if len(answers) >= 3 else ""
    for index, line in enumerate(group_lines):
        if index == 0:
            result.extend(wenyan.styled_body_lines([wenyan.strip_fillable_blanks_from_line(line)], section_title))
            if result:
                meta = dict(result[-1].meta or {})
                meta["force_new_chunk"] = True
                result[-1].meta = meta
            continue
        value = re.sub(r'<span\s+class="blank"[^>]*>\s*</span>', "", line)
        stripped = re.sub(r"[_＿]+", "", plain(value)).strip()
        if "句式类型" in stripped and "句意" in stripped:
            sentence = re.split(r"句式类型", stripped, maxsplit=1)[0].strip()
            result.extend(
                wenyan.styled_body_lines(
                    [f"{sentence}　　句式类型：（{wenyan.answer_span(sentence_type)}）　句意："],
                    section_title,
                )
            )
            result.append(wenyan.StyledLine(sanbu_fixed_answer_line(meaning), "long_answer"))
            continue
        if stripped.startswith("请回忆学过的课文"):
            prefix = re.sub(r"[:：]\s*$", "：", stripped)
            result.extend(
                wenyan.styled_body_lines(
                    [f"{prefix}{wenyan.answer_span(same_sentence, underline=True)}"],
                    section_title,
                )
            )
            continue
        result.extend(wenyan.styled_body_lines([normalize_sentence_pattern_line(line, item)], section_title))
    if item.analyses:
        result.append(wenyan.StyledLine("", "answer_spacer"))
        for analysis in item.analyses:
            result.append(wenyan.StyledLine(analysis, "analysis"))
    return result


def build_sanbu_glossary_reading_section(
    lines: list[str],
    answer_lines: list[str] | None,
) -> tuple[list[wenyan.StyledLine], bool]:
    section_title = lines[0] if lines else ""
    answer_items = wenyan.extract_answer_items(answer_lines)
    leading, groups = wenyan.split_body_numbered_groups(lines)
    styled: list[wenyan.StyledLine] = []

    preamble_lines = build_sanbu_glossary_preamble_lines(leading, answer_lines)
    styled.extend(wenyan.styled_body_lines(preamble_lines, section_title))
    matched_any = bool(parse_sanbu_glossary_entries(answer_lines))

    for num, group_lines in groups:
        item = answer_items.get(num)
        if item is not None and is_sanbu_translation_group(group_lines):
            styled.extend(format_sanbu_translation_group(group_lines, item, section_title))
            matched_any = True
            continue
        if item is not None and is_sentence_pattern_group(group_lines):
            styled.extend(format_sentence_pattern_group(group_lines, item, section_title))
            matched_any = True
            continue
        partial_styled, partial_matched = wenyan.build_styled_section(group_lines, answer_lines)
        styled.extend(partial_styled)
        matched_any = matched_any or partial_matched
    return styled, matched_any


def build_sanbu_styled_section(
    lines: list[str],
    answer_lines: list[str] | None,
) -> tuple[list[wenyan.StyledLine], bool]:
    lines = mark_reading_passage_underlines_blue(lines)
    if is_sanbu_glossary_reading(lines, answer_lines):
        return build_sanbu_glossary_reading_section(lines, answer_lines)
    lines = [normalize_sanbu_fillable_blanks(line) for line in lines]
    section_title = lines[0] if lines else ""
    answer_items = wenyan.extract_answer_items(answer_lines)
    leading, groups = wenyan.split_body_numbered_groups(lines)
    if not (answer_items and groups):
        return wenyan.build_styled_section(lines, answer_lines)

    styled: list[wenyan.StyledLine] = []
    preamble_lines, preamble_filled = wenyan.fill_lines_with_answer_values(
        leading,
        wenyan.extract_preamble_answers(answer_lines),
    )
    styled.extend(wenyan.styled_body_lines(preamble_lines, section_title))
    matched_any = preamble_filled
    for num, group_lines in groups:
        item = answer_items.get(num)
        if item is not None and is_sanbu_translation_group(group_lines):
            styled.extend(format_sanbu_translation_group(group_lines, item, section_title))
            matched_any = True
            continue
        if item is not None and is_sentence_pattern_group(group_lines):
            styled.extend(format_sentence_pattern_group(group_lines, item, section_title))
            matched_any = True
            continue
        partial_styled, partial_matched = wenyan.build_styled_section(group_lines, answer_lines)
        # The generic builder sees this one-question list as its own section; drop
        # repeated leading noise and only keep the formatted question block.
        styled.extend(partial_styled)
        matched_any = matched_any or partial_matched
    return styled, matched_any


def install_sanbu_hooks() -> None:
    global _SANBU_HOOKS_INSTALLED
    if _SANBU_HOOKS_INSTALLED:
        return
    old_is_long_answer_prompt = wenyan.is_long_answer_prompt
    old_extract_ordered_answers = wenyan.extract_ordered_answers
    old_role_style = wenyan.role_style
    old_add_content = wenyan.add_content
    old_append_title_volume_label = wenyan.append_title_volume_label

    def sanbu_is_long_answer_prompt(section_title: str, group_lines: list[str]) -> bool:
        if old_is_long_answer_prompt(section_title, group_lines):
            return True
        plain_group = "\n".join(plain(line) for line in group_lines)
        triggers = [
            "请引用",
            "请根据",
            "进行回答",
            "原因",
            "依据",
            "品质",
            "体现",
            "怎样",
            "如何",
            "为什么",
            "有何不同",
            "侧重点",
            "主张",
            "看法",
        ]
        return any(trigger in plain_group for trigger in triggers)

    def sanbu_extract_ordered_answers(answer_lines: list[str] | None) -> list[str]:
        if not answer_lines:
            return []
        filtered: list[str] = []
        for line in answer_lines:
            value = plain(line)
            if value in {"【重点字词】", "【词语积累】"}:
                continue
            filtered.append(line)
        return old_extract_ordered_answers(filtered)

    def sanbu_extract_answer_items(answer_lines: list[str] | None) -> dict[int, wenyan.AnswerItem]:
        if not answer_lines:
            return {}
        items: dict[int, wenyan.AnswerItem] = {}
        current: wenyan.AnswerItem | None = None
        skipping_reference_tail = False
        for raw_line in answer_lines:
            line = raw_line.strip()
            if not line or line in wenyan.MODULE_TITLES or line in wenyan.VOLUMES or wenyan.is_section_marker(line):
                skipping_reference_tail = False
                continue
            if line in {"【重点字词】", "【词语积累】"}:
                continue
            match = wenyan.answer_item_match(line)
            if match:
                skipping_reference_tail = False
                current = wenyan.AnswerItem(answers=[], analyses=[])
                items[int(match.group(1))] = current
                append_sanbu_answer_item_text(current, match.group(2), old_extract_ordered_answers)
                continue
            if current is not None and SANBU_ANALYSIS_LABEL_RE.match(line):
                current.analyses.append(line)
                skipping_reference_tail = False
                continue
            if skipping_reference_tail:
                continue
            if wenyan.looks_like_reference_translation_tail(line, current):
                skipping_reference_tail = True
                continue
            if current is not None:
                append_sanbu_answer_item_text(current, line, old_extract_ordered_answers)
        return items

    def sanbu_make_piece_slide_chunks(
        body_piece: wenyan.Piece,
        answer_piece: wenyan.Piece | None,
    ) -> list[list[wenyan.StyledLine]]:
        body_sections = split_sanbu_piece_lines(body_piece)
        answer_sections = split_sanbu_piece_lines(answer_piece)
        meta = sanbu_meta_info(body_piece)
        chunks: list[list[wenyan.StyledLine]] = []
        for section_name in ("main", "reference", "vocab"):
            body_lines = body_sections.get(section_name, [])
            answer_lines = answer_sections.get(section_name)
            if section_name != "reference" and not body_lines:
                continue
            if section_name == "reference":
                if not answer_lines:
                    continue
                reference_text = [
                    line
                    for line in answer_lines
                    if plain(line) and "参考译文" not in plain(line) and "<img" not in line
                ]
                if not reference_text:
                    continue
                styled = [wenyan.StyledLine("参考译文", "reference_header")] + [
                    wenyan.StyledLine(line, "reference_translation") for line in reference_text
                ]
                chunks.extend(wenyan.chunk_styled_lines(styled, max_units=9))
                continue

            if section_name == "vocab":
                styled, _matched = build_sanbu_vocab_styled_section(body_lines, answer_lines)
            else:
                styled, _matched = build_sanbu_styled_section(body_lines, answer_lines)
                styled = add_slash_answer_labels(styled)
            max_units = 12
            if section_name == "main" and meta:
                styled = [wenyan.StyledLine("", "sanbu_meta", meta=meta)] + styled
                chunks.extend(chunk_sanbu_meta_main_lines(styled))
                continue
            if section_name == "vocab":
                styled = [wenyan.StyledLine("词语积累", "heading")] + [
                    line for line in styled if plain(line.text) != "词语积累"
                ]
            chunks.extend(repair_orphan_sanbu_question_chunks(wenyan.chunk_styled_lines(styled, max_units=max_units)))
        return chunks

    def sanbu_add_cover_content(root, volume: str, start: int, end: int) -> None:
        display_volume = volume.replace(" 单文本阅读", "").replace(" 多文本阅读", "")
        piece_range = f"第{start}-{end}篇"
        wenyan.replace_exact_text(root, "必修上册", display_volume)
        wenyan.replace_exact_text(root, "篇目1-4", piece_range)
        for sp in wenyan.text_shapes(root):
            text = wenyan.get_shape_text(sp).strip()
            if "文言梳理" in text:
                wenyan.set_shape_text(sp, "高中文言文三步进阶")
                wenyan.style_existing_text_shape(sp, font="黑体", color="000000", bold=True)
                continue
            if re.fullmatch(r"篇目\d+-\d+", text):
                wenyan.set_shape_text(sp, piece_range)
                text = piece_range
            if text in {
                display_volume,
                piece_range,
                f"{display_volume} {piece_range}",
                f"{display_volume}{piece_range}",
            }:
                wenyan.style_existing_text_shape(sp, font="宋体", color="000000", bold=False)
                set_text_shape_size(sp, 26)
                if text == display_volume:
                    set_shape_box(sp, x=6.35, y=5.47, w=3.35, h=0.58)
                elif text == piece_range:
                    set_shape_box(sp, x=6.35, y=6.08, w=3.35, h=0.58)

    def sanbu_append_title_volume_label(root, volume: str) -> None:
        display_volume = volume.replace(" 单文本阅读", "").replace(" 多文本阅读", "")
        if display_volume in {"基础篇30篇", "提升篇20篇", "进阶篇15篇"}:
            wenyan.append_textbox(
                root,
                x=0.38,
                y=6.42,
                w=4.20,
                h=0.70,
                lines=[wenyan.StyledLine(wenyan.spaced_volume_label(display_volume), "title_volume")],
                name="篇目标题册别",
            )
            return
        old_append_title_volume_label(root, volume)

    def sanbu_role_style(role: str) -> tuple[str, float, str, bool]:
        if role == "sanbu_meta_top":
            return old_role_style(role)[0], 24, "000000", False
        if role == "sanbu_meta_info":
            return old_role_style(role)[0], 24, "000000", False
        if role == "sanbu_meta_top":
            return "宋体", 22, "000000", False
        if role == "sanbu_meta_info":
            return "宋体", 20, "000000", False
        if role == "reference_header":
            return "宋体", 18, "000000", False
        if role == "reference_translation":
            return "KaiTi", 24, "FF0000", False
        return old_role_style(role)

    def sanbu_add_content(
        root,
        rels_root,
        piece: wenyan.Piece,
        chunk: list[wenyan.StyledLine],
        *,
        media_assets: dict[str, bytes],
        tif_search_dirs: list[Path],
        content_box: tuple[float, float, float, float],
    ) -> None:
        meta_lines = [line for line in chunk if line.role == "sanbu_meta"]
        if not meta_lines:
            if chunk and chunk[0].role == "reference_header":
                wenyan.clear_text_shapes(root, keep_substrings=("能力",))
                base_x, base_y, base_w, base_h = content_box
                tif_path = wenyan.resolve_tif_path("参考译文.TIF", tif_search_dirs) or wenyan.resolve_tif_path("参考译文.tif", tif_search_dirs)
                if tif_path is not None:
                    wenyan.append_tif_picture(
                        root,
                        rels_root,
                        media_assets,
                        tif_path=tif_path,
                        x=base_x,
                        y=base_y + 0.08,
                        h=0.34,
                        name="参考译文",
                    )
                content_lines = [line for line in chunk if line.role != "reference_header"]
                y = base_y + 0.64
                h = max(0.4, base_h - 0.64)
                wenyan.append_textbox(
                    root,
                    x=base_x,
                    y=y,
                    w=base_w,
                    h=h,
                    lines=content_lines,
                    name=f"{piece.title} reference translation",
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
                wenyan.set_slide_animations(root, [])
                return
            if any(line.meta and line.meta.get("l_icon") for line in chunk):
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
                wenyan.append_tif_overlays(
                    root,
                    rels_root,
                    media_assets,
                    x=base_x,
                    y=base_y,
                    w=base_w,
                    lines=chunk,
                    tif_search_dirs=tif_search_dirs,
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
                return
            if any((line.role or "") == "long_answer" for line in chunk):
                wenyan.clear_text_shapes(root, keep_substrings=("能力",))
                base_x, base_y, base_w, base_h = content_box
                render_sanbu_text_content(
                    root,
                    rels_root,
                    media_assets,
                    piece=piece,
                    lines=chunk,
                    tif_search_dirs=tif_search_dirs,
                    x=base_x,
                    y=base_y,
                    w=base_w,
                    h=base_h,
                    name_suffix="content",
                )
                return
            old_add_content(
                root,
                rels_root,
                piece,
                chunk,
                media_assets=media_assets,
                tif_search_dirs=tif_search_dirs,
                content_box=content_box,
            )
            return

        meta = meta_lines[0].meta if isinstance(meta_lines[0].meta, dict) else {}
        content_lines = [line for line in chunk if line.role != "sanbu_meta"]
        wenyan.clear_text_shapes(root, keep_substrings=("能力",))
        append_sanbu_meta_header(root, rels_root, media_assets, {str(k): str(v) for k, v in meta.items()})

        base_x, base_y, base_w, base_h = content_box
        y = max(base_y, 2.76)
        h = max(0.4, base_h - (y - base_y))
        render_sanbu_text_content(
            root,
            rels_root,
            media_assets,
            piece=piece,
            lines=content_lines,
            tif_search_dirs=tif_search_dirs,
            x=base_x,
            y=y,
            w=base_w,
            h=h,
            name_suffix="content",
        )

    wenyan.split_analysis = sanbu_split_analysis
    wenyan.is_long_answer_prompt = sanbu_is_long_answer_prompt
    wenyan.extract_ordered_answers = sanbu_extract_ordered_answers
    wenyan.extract_answer_items = sanbu_extract_answer_items
    wenyan.fixed_long_answer_lines = sanbu_fixed_long_answer_lines
    wenyan.make_piece_slide_chunks = sanbu_make_piece_slide_chunks
    wenyan.add_cover_content = sanbu_add_cover_content
    wenyan.append_title_volume_label = sanbu_append_title_volume_label
    wenyan.role_style = sanbu_role_style
    wenyan.add_content = sanbu_add_content
    _SANBU_HOOKS_INSTALLED = True


def find_md_files(run_dir: Path) -> tuple[Path, Path]:
    body = sorted(run_dir.glob("正文_*.md"))
    answer = sorted(run_dir.glob("答案_*.md"))
    if not body or not answer:
        raise FileNotFoundError(f"未在目录中找到正文_*.md 和 答案_*.md：{run_dir}")
    return body[0], answer[0]


def sanbu_group_piece_number(piece: wenyan.Piece, group_volume: str) -> int | None:
    if group_volume != "进阶篇15篇":
        return piece.num if piece.volume == group_volume else None
    if piece.volume == "进阶篇15篇 单文本阅读":
        return piece.num
    if piece.volume == "进阶篇15篇 多文本阅读":
        return piece.num + 4
    return None


def sanbu_display_piece(piece: wenyan.Piece, display_num: int) -> wenyan.Piece:
    if display_num == piece.num:
        return piece
    return wenyan.Piece(
        volume=piece.volume,
        num=display_num,
        title=piece.title,
        author=piece.author,
        lines=piece.lines,
    )


@isolated_generation
def generate_sanbu_pptx(
    *,
    md_run: Path,
    template: Path,
    output_dir: Path,
    limit_groups: int | None = 2,
    group_ranges: list[dict[str, object]] | None = None,
    progress: Callable[[str], None] | None = None,
) -> list[Path]:
    def emit(message: str) -> None:
        if progress:
            progress(message)
        else:
            print(message)

    install_sanbu_hooks()
    body_md, answer_md = find_md_files(md_run)
    emit(f"正文 MD：{body_md}")
    emit(f"答案 MD：{answer_md}")
    emit(f"模板 PPTX：{template}")

    body_pieces = parse_body_pieces(body_md)
    answer_pieces = parse_answer_pieces(answer_md, body_pieces)
    answer_map = {wenyan.piece_key(piece): piece for piece in answer_pieces}
    tif_search_dirs = wenyan.infer_tif_search_dirs(md_run)
    if SANBU_SOURCE_DIR.exists() and SANBU_SOURCE_DIR not in tif_search_dirs:
        tif_search_dirs.append(SANBU_SOURCE_DIR)
    selected_groups = SANBU_GROUPS if limit_groups is None else SANBU_GROUPS[:limit_groups]
    if group_ranges is not None:
        selected_groups = []
        for index, group in enumerate(group_ranges, 1):
            volume = str(group["volume"])
            start = int(group["start"])
            end = int(group["end"])
            keys = {tuple(key) for key in group["unit_keys"]}  # type: ignore[arg-type]
            selected = [piece for piece in body_pieces if (piece.volume, piece.num) in keys]
            selected_groups.append((f"{index}{volume} 第{start}-{end}篇.pptx", volume, start, end, selected))
    else:
        selected_groups = [(*group, None) for group in selected_groups]

    emit(f"正文篇目数：{len(body_pieces)}")
    emit(f"答案篇目数：{len(answer_pieces)}")
    output_dir.mkdir(parents=True, exist_ok=True)
    generated: list[Path] = []
    for group_index, (file_name, volume, start, end, selected_override) in enumerate(selected_groups, 1):
        group_pairs: list[tuple[wenyan.Piece, wenyan.Piece]] = []
        if selected_override is not None:
            for piece in selected_override:
                display_num = sanbu_group_piece_number(piece, volume)
                group_pairs.append((piece, sanbu_display_piece(piece, display_num or piece.num)))
        else:
            for piece in body_pieces:
                display_num = sanbu_group_piece_number(piece, volume)
                if display_num is not None and start <= display_num <= end:
                    group_pairs.append((piece, sanbu_display_piece(piece, display_num)))
        group_pieces = [display_piece for _source_piece, display_piece in group_pairs]
        if not group_pieces:
            emit(f"跳过，无篇目：{file_name}")
            continue
        deck_answer_map = dict(answer_map)
        for source_piece, display_piece in group_pairs:
            answer_piece = answer_map.get(wenyan.piece_key(source_piece))
            if answer_piece is not None:
                deck_answer_map[wenyan.piece_key(display_piece)] = answer_piece
        output = output_dir / file_name
        emit(f"[{group_index}/{len(selected_groups)}] 生成：{output.name}")
        wenyan.build_deck(template, output, group_pieces, deck_answer_map, volume, start, end, tif_search_dirs=tif_search_dirs)
        generated.append(output)
        emit(f"已生成：{output}，篇目：{', '.join(f'{p.num}.{p.title}' for p in group_pieces)}")

    manifest = {
        "md_run": str(md_run),
        "template": str(template),
        "generated": [str(path) for path in generated],
    }
    (output_dir / "sanbu_test_generation_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return generated


def scan_sanbu_group_units(md_run: Path) -> list[dict[str, object]]:
    """Return detected 三步进阶 sections and passages for GUI grouping."""

    body_md, _answer_md = find_md_files(md_run)
    return [
        {"volume": piece.volume, "num": piece.num, "title": piece.title}
        for piece in parse_body_pieces(body_md)
    ]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="测试：三步进阶 md + 模板生成 PPTX。")
    parser.add_argument("--md-run", default=str(DEFAULT_MD_RUN), help="三步进阶 md run 目录。")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="PPTX 模板路径。")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="测试 PPTX 输出目录。")
    parser.add_argument("--limit-groups", type=int, default=2, help="生成前几个分组；默认 2，传 0 表示全部。")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    limit_groups = None if args.limit_groups == 0 else args.limit_groups
    generate_sanbu_pptx(
        md_run=Path(args.md_run),
        template=Path(args.template),
        output_dir=Path(args.output_dir),
        limit_groups=limit_groups,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

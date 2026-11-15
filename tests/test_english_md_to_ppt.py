#!/usr/bin/env python3
"""Experimental Markdown-to-PPTX generator for 高中英语词汇语境记忆.

This is a test script only. It stays outside the GUI/main program while the
English book rules are being learned and adjusted.
"""

from __future__ import annotations

import argparse
import html
import json
import re
import shutil
import sys
import zipfile
from dataclasses import dataclass
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path
from typing import Callable, Iterable
from xml.etree import ElementTree as ET


PROJECT_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
TESTS_DIR = Path(__file__).resolve().parent
if str(TESTS_DIR) not in sys.path:
    sys.path.insert(0, str(TESTS_DIR))

import test_md_to_ppt as wenyan  # noqa: E402


DEFAULT_MD_RUN = PROJECT_DIR / "out_md" / "英语" / "run_20260711_213151"
DEFAULT_TEMPLATE = PROJECT_DIR / "高中语境记忆模板.pptx"
DEFAULT_OUTPUT_DIR = PROJECT_DIR / "output" / "test_english_ppt"

BOOK_TITLE = "新高中英语词汇语境记忆"
EXERCISE_TITLES = [
    "目标词汇",
    "语法填空",
    "核对译文",
    "重点词汇拓展",
    "词块填空",
    "句型突破",
    "写作演练",
]
EXERCISE_NUMBERS = {
    "目标词汇": "01",
    "语法填空": "02",
    "核对译文": "03",
    "重点词汇拓展": "04",
    "词块填空": "05",
    "句型突破": "06",
    "写作演练": "07",
}
EXERCISE_TO_ANSWER = {
    "目标词汇": 1,
    "语法填空": 2,
    "重点词汇拓展": 4,
    "词块填空": 5,
    "句型突破": 6,
    "写作演练": 7,
}

BLUE = "3F587F"
ORANGE = "ED7D31"
BLACK = "000000"
WHITE = "FFFFFF"
LATIN_FONT = "Times New Roman"
CJK_FONT = "Microsoft YaHei"
CONTENT_LINE_SPACING_PT = 32.0
GRAMMAR_FIRST_LINE_INDENT_IN = 0.56
CONTENT_BOTTOM_Y = 6.62
POS_ABBR = r"(?:adv|adj|n|v|vt|vi|prep|pron|conj|phr|num|art|interj)"
POS_TOKEN_RE = re.compile(rf"(?<![A-Za-z])({POS_ABBR}\.)", flags=re.I)


@dataclass
class HtmlRun:
    text: str
    underline: bool = False
    answer: bool = False
    bold: bool = False
    italic: bool = False


@dataclass
class RichLine:
    text: str
    role: str = "body"


@dataclass
class BuiltSection:
    title: str
    chunks: list[list[RichLine]]
    title_slide_number: int


@dataclass
class Passage:
    num: int
    theme: str
    sections: dict[str, list[str]]


@dataclass
class PassageAnswers:
    exercises: dict[int, list[str]]


class SimpleInlineParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.runs: list[HtmlRun] = []
        self.stack: list[tuple[bool, bool, bool, bool]] = []
        self.underline = False
        self.answer = False
        self.bold = False
        self.italic = False

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self.stack.append((self.underline, self.answer, self.bold, self.italic))
        attr = dict(attrs)
        classes = set((attr.get("class") or "").split())
        style = attr.get("style") or ""
        if tag in {"b", "strong"}:
            self.bold = True
        if tag in {"i", "em"}:
            self.italic = True
        if "bold" in classes or re.search(r"font-weight\s*:\s*(bold|[6-9]00)", style, flags=re.I):
            self.bold = True
        if tag == "span":
            if "uline" in classes:
                self.underline = True
            if "answer-fill" in classes:
                self.answer = True
                self.bold = True

    def handle_endtag(self, tag: str) -> None:
        if self.stack:
            self.underline, self.answer, self.bold, self.italic = self.stack.pop()

    def handle_data(self, data: str) -> None:
        if data:
            self.runs.append(HtmlRun(data, self.underline, self.answer, self.bold, self.italic))


def strip_style_block(md: str) -> str:
    return re.sub(r"(?is)<style.*?</style>", "", md).strip()


def blank_underscores(style: str) -> str:
    match = re.search(r"width:\s*([0-9.]+)em", style)
    width = float(match.group(1)) if match else 3.0
    return "_" * max(4, min(34, int(round(width * 2.4))))


def normalize_line(line: str) -> str:
    return html.unescape(line.strip())


def clean_lines(md: str) -> list[str]:
    text = strip_style_block(md)
    return [normalize_line(line) for line in text.splitlines()]


def plain_text(line: str) -> str:
    line = re.sub(r'<span\s+class="blank"([^>]*)></span>', lambda m: blank_underscores(m.group(1) or ""), line)
    line = re.sub(r"<[^>]+>", "", line)
    return html.unescape(line).strip()


def answer_plain_text(line: str) -> str:
    line = re.sub(r'<span\s+class="blank"[^>]*></span>', " ", line)
    line = re.sub(r"<[^>]+>", "", line)
    return html.unescape(line).strip()


def is_noise_line(line: str) -> bool:
    text = plain_text(line)
    if not text:
        return True
    if text == BOOK_TITLE:
        return True
    if text in {"Passage", "Exercise[]", "[]"}:
        return True
    if re.fullmatch(r"[\d\s*×xX.]+", text):
        return True
    if set(text) <= {"_"}:
        return True
    return False


def theme_for_passage(num: int) -> str:
    if num <= 20:
        return "人与自我"
    if num <= 40:
        return "人与社会"
    return "人与自然"


def parse_body_passages(path: Path) -> list[Passage]:
    lines = clean_lines(path.read_text(encoding="utf-8"))
    starts = [i for i, line in enumerate(lines) if plain_text(line) == "目标词汇"]
    passages: list[Passage] = []
    for passage_index, start in enumerate(starts, 1):
        end = starts[passage_index] if passage_index < len(starts) else len(lines)
        chunk = lines[start:end]
        sections: dict[str, list[str]] = {}
        current: str | None = None
        for raw in chunk:
            title = plain_text(raw)
            if title in EXERCISE_TITLES:
                current = title
                sections[current] = []
                continue
            keep_section_blank = current in {"句型突破", "重点词汇拓展"} and '<span class="blank"' in raw
            if current is None or (is_noise_line(raw) and not keep_section_blank):
                continue
            sections[current].append(raw)
        passages.append(Passage(passage_index, theme_for_passage(passage_index), sections))
    return passages


def parse_answer_passages(path: Path) -> dict[int, PassageAnswers]:
    lines = clean_lines(path.read_text(encoding="utf-8"))
    answers: dict[int, PassageAnswers] = {}
    current_passage: int | None = None
    current_exercise: int | None = None
    for raw in lines:
        text = plain_text(raw)
        if not text:
            continue
        passage_match = re.search(r"Passage\s+(\d+)", text, flags=re.I)
        if passage_match:
            current_passage = int(passage_match.group(1))
            answers[current_passage] = PassageAnswers({})
            current_exercise = None
            continue
        exercise_match = re.match(r"Exercise\s+(\d+)", text, flags=re.I)
        if exercise_match and current_passage is not None:
            current_exercise = int(exercise_match.group(1))
            answers[current_passage].exercises[current_exercise] = []
            continue
        if current_passage is not None and current_exercise is not None:
            answers[current_passage].exercises[current_exercise].append(raw)
    return answers


def parse_booklet_vocab_passages(path: Path | None) -> dict[int, list[str]]:
    if not path or not path.exists():
        return {}
    lines = clean_lines(path.read_text(encoding="utf-8"))
    passages: dict[int, list[str]] = {}
    current_passage: int | None = None
    for raw in lines:
        text = plain_text(raw)
        if not text:
            continue
        passage_match = re.search(r"Passage\s+(\d+)", text, flags=re.I)
        if passage_match:
            current_passage = int(passage_match.group(1))
            passages.setdefault(current_passage, [])
            continue
        if current_passage is None:
            continue
        if re.match(r"^\d+\.", text):
            passages[current_passage].append(text)
    return passages


def parse_target_vocab_parts(line: str) -> tuple[str, str, str | None, str] | None:
    text = plain_text(line)
    match = re.match(r"^(\d+)\.(.*?)(?:_+)?$", text)
    if not match:
        return None
    num, body = match.groups()
    body = re.sub(r"_+$", "", body).strip()
    pos_match = POS_TOKEN_RE.search(body)
    if pos_match:
        word = body[: pos_match.start()].strip()
        pos = pos_match.group(1)
        tail = body[pos_match.end() :].strip()
    else:
        word = body.strip()
        pos = None
        tail = ""
    return num, word, pos.lower() if pos else None, tail


def parse_target_vocab_prompt(line: str) -> tuple[str, str | None] | None:
    parts = parse_target_vocab_parts(line)
    if not parts:
        return None
    return parts[1], parts[2]


def strip_booklet_phonetic(text: str) -> str:
    text = re.sub(r"［[^］]*］", " ", text)
    text = re.sub(r"\[[^\]]*\]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def extract_meaning_for_pos(text: str, pos: str | None) -> str:
    if not pos:
        return text.strip()
    matches = list(POS_TOKEN_RE.finditer(text))
    if not matches:
        return text.strip()
    wanted = pos.lower()
    for index, match in enumerate(matches):
        if match.group(1).lower() != wanted:
            continue
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        return text[start:end].strip(" ；;，,")
    # If the booklet has only one part of speech but the prompt POS differs
    # slightly, still prefer the definition body over returning the whole tail.
    if len(matches) == 1:
        return text[matches[0].end() :].strip(" ；;，,")
    return text.strip()


def booklet_vocab_answer(entries: list[str], word: str, pos: str | None) -> str | None:
    wanted = re.sub(r"\s+", " ", word).strip().lower()
    for entry in entries:
        match = re.match(r"^\d+\.(.*)$", strip_booklet_phonetic(entry))
        if not match:
            continue
        body = match.group(1).strip()
        normalized_body = re.sub(r"\s+", " ", body).strip()
        if not normalized_body.lower().startswith(wanted):
            continue
        tail = normalized_body[len(word) :].strip()
        if not tail:
            continue
        answer = extract_meaning_for_pos(tail, pos)
        return answer or None
    return None


def split_semicolon_answers(lines: Iterable[str]) -> list[str]:
    text = " ".join(answer_plain_text(line) for line in lines)
    parts = re.split(r"[;；]\s*", text)
    return [part.strip() for part in parts if part.strip()]


def parse_numbered_answers(lines: Iterable[str]) -> dict[int, str]:
    text = "\n".join(answer_plain_text(line) for line in lines)
    text = re.sub(r"\s+", " ", text)
    result: dict[int, str] = {}
    for match in re.finditer(r"(\d+)\.(.*?)(?=\s+\d+\.|$)", text):
        value = match.group(2).strip()
        if value:
            result[int(match.group(1))] = value
    return result


def parse_exercise4_answers(lines: list[str]) -> tuple[dict[int, list[str]], dict[int, str]]:
    text = "\n".join(answer_plain_text(line) for line in lines)
    text = re.sub(r"\s+", " ", text)
    b_pos = text.find(" b.")
    if b_pos < 0:
        b_pos = text.find("b.")
    a_text = text
    b_text = ""
    if b_pos >= 0:
        a_text = text[:b_pos]
        b_text = text[b_pos:]
    a_text = re.sub(r"^a\.", "", a_text).strip()
    b_text = re.sub(r"^b\.", "", b_text).strip()
    a_result: dict[int, list[str]] = {}
    for match in re.finditer(r"(\d+)\.(.*?)(?=\s+\d+\.|$)", a_text):
        values = [item.strip() for item in re.split(r"[;；]", match.group(2)) if item.strip()]
        if values:
            a_result[int(match.group(1))] = values
    b_result = parse_numbered_answers([b_text]) if b_text else {}
    return a_result, b_result


def parse_parenthetical_answers(lines: Iterable[str]) -> list[str]:
    values: list[str] = []
    for line in lines:
        text = answer_plain_text(line)
        if re.match(r"^[（(]\d+[）)]", text):
            values.append(re.sub(r"^[（(]\d+[）)]\s*", "", text).strip())
    return values


def replace_blank_spans(line: str, answers: list[str] | None = None, *, underline_answer: bool = True) -> str:
    index = 0

    def repl(match: re.Match[str]) -> str:
        nonlocal index
        style = match.group(1) or ""
        if answers is not None and index < len(answers):
            value = html.escape(answers[index], quote=False)
            index += 1
            cls = "answer-fill uline" if underline_answer else "answer-fill"
            return f'<span class="{cls}">{value}</span>'
        return blank_underscores(style)

    line = re.sub(r'<span\s+class="blank"([^>]*)></span>', repl, line)
    return line


def replace_blank_spans_with_spaces(line: str, spaces: int = 8) -> str:
    return re.sub(r'<span\s+class="blank"[^>]*></span>', " " * spaces, line)


def fill_lines_with_answers(lines: list[str], answers: list[str]) -> list[str]:
    answer_index = 0
    filled: list[str] = []
    for line in lines:
        blank_count = len(re.findall(r'<span\s+class="blank"', line))
        if blank_count and answer_index < len(answers):
            line_answers = answers[answer_index : answer_index + blank_count]
            filled.append(replace_blank_spans(line, line_answers))
            answer_index += blank_count
        else:
            filled.append(replace_blank_spans(line))
    return filled


def fill_numbered_blank_lines(lines: list[str], answer_values: list[str]) -> list[str]:
    answer_index = 0
    out: list[str] = []
    waiting_for_answer = False
    for line in lines:
        text = plain_text(line)
        if re.match(r"^[（(]\d+[）)]", text):
            out.append(line)
            waiting_for_answer = True
            continue
        if set(text) <= {"_"}:
            if waiting_for_answer and answer_index < len(answer_values):
                answer = html.escape(answer_values[answer_index], quote=False)
                out.append(f'<span class="answer-fill uline">{answer}</span>')
                answer_index += 1
            waiting_for_answer = False
            continue
        waiting_for_answer = False
        if '<span class="blank"' in line and answer_index < len(answer_values):
            blank_count = len(re.findall(r'<span\s+class="blank"', line))
            values = answer_values[answer_index : answer_index + blank_count]
            out.append(replace_blank_spans(line, values))
            answer_index += blank_count
        else:
            out.append(replace_blank_spans(line))
    return out


def format_vocab_line(line: str, answer: str | None = None) -> str:
    parts = parse_target_vocab_parts(line)
    if not parts:
        return replace_blank_spans(line)
    num, word, pos, tail = parts
    if pos:
        head = f"{num}.{html.escape(word, quote=False)} <i>{html.escape(pos, quote=False)}</i>"
        if tail:
            head += f" {html.escape(tail, quote=False)}"
    else:
        head = f"{num}.{html.escape(word, quote=False)}"
        if tail:
            head += f" {html.escape(tail, quote=False)}"
    if answer:
        return f'{head}  <span class="answer-fill">{html.escape(answer, quote=False)}</span>'
    return f"{head}  ________________"


def extract_vocab_terms(lines: list[str]) -> list[str]:
    terms: list[str] = []
    for line in lines:
        text = plain_text(line)
        match = re.match(r"^\d+\.(.*?)(?:_+)?$", text)
        if not match:
            continue
        body = match.group(1).strip()
        pos_match = re.match(rf"^(.+?)\s+({POS_ABBR}\.)\s*$", body, flags=re.I)
        if pos_match:
            term = pos_match.group(1).strip()
        else:
            term = body.strip()
        if term and term.lower() not in {item.lower() for item in terms}:
            terms.append(term)
    return sorted(terms, key=len, reverse=True)


def emphasize_terms(line: str, terms: list[str]) -> str:
    if not terms:
        return line
    pieces = re.split(r"(<[^>]+>)", line)
    out: list[str] = []
    for piece in pieces:
        if piece.startswith("<") and piece.endswith(">"):
            out.append(piece)
            continue
        rendered = piece
        for term in terms:
            escaped = re.escape(term)
            pattern = rf"(?<![A-Za-z])({escaped})(?![A-Za-z])"
            rendered = re.sub(pattern, r"<b>\1</b>", rendered, flags=re.I)
        out.append(rendered)
    return "".join(out)


def target_vocab_lines(lines: list[str], answer_lines: list[str], booklet_entries: list[str] | None = None) -> list[RichLine]:
    answers = split_semicolon_answers(answer_lines)
    result: list[RichLine] = []
    answer_index = 0
    for line in lines:
        text = plain_text(line)
        if not re.match(r"^\d+\.", text):
            continue
        answer = None
        prompt = parse_target_vocab_prompt(line)
        if booklet_entries and prompt:
            answer = booklet_vocab_answer(booklet_entries, prompt[0], prompt[1])
        if answer is None and answer_index < len(answers):
            answer = answers[answer_index]
        if answer:
            answer_index += 1
        result.append(RichLine(format_vocab_line(line, answer), "vocab"))
        if len(result) >= 12:
            break
    return result


def grammar_lines(lines: list[str], answer_lines: list[str], vocab_terms: list[str]) -> list[RichLine]:
    answers = split_semicolon_answers(answer_lines)
    filled = fill_lines_with_answers(lines, answers)
    result: list[RichLine] = []
    for line in filled:
        text = plain_text(line)
        if not text:
            continue
        if len(result) == 0:
            result.append(RichLine(line, "article_title"))
        else:
            result.append(RichLine(emphasize_terms(line, vocab_terms), "article"))
    return result


def translation_lines(lines: list[str]) -> list[RichLine]:
    result: list[RichLine] = []
    for index, line in enumerate(lines):
        text = plain_text(line)
        if not text:
            continue
        result.append(RichLine(line, "zh_title" if index == 0 else "translation"))
    return result


def expansion_lines(lines: list[str], answer_lines: list[str]) -> list[RichLine]:
    a_answers, b_answers = parse_exercise4_answers(answer_lines)
    result: list[RichLine] = []
    in_b = False
    for line in lines:
        text = plain_text(line)
        if not text:
            continue
        if text.startswith("a."):
            result.append(RichLine(text, "exercise_head"))
            continue
        if text.startswith("b."):
            in_b = True
            result.append(RichLine(text, "exercise_head"))
            continue
        if in_b and set(text) <= {"_"}:
            continue
        num_match = re.match(r"^(\d+)\.", text)
        if num_match:
            num = int(num_match.group(1))
            if in_b:
                answer = b_answers.get(num)
                rendered = replace_blank_spans(line)
                if answer:
                    rendered = (
                        f'{rendered} <span class="answer-fill uline">'
                        f'{html.escape(answer, quote=False)}</span>'
                    )
                result.append(RichLine(rendered, "body"))
            else:
                rendered = fill_lines_with_answers([line], a_answers.get(num, []))[0]
                result.append(RichLine(rendered, "body"))
            continue
        result.append(RichLine(replace_blank_spans(line), "body"))
    return result


def phrase_lines(lines: list[str], answer_lines: list[str]) -> list[RichLine]:
    numbered = parse_numbered_answers(answer_lines)
    values = [numbered[i] for i in sorted(numbered)]
    bank_lines: list[str] = []
    question_lines: list[str] = []
    seen_question = False
    for line in lines:
        if re.match(r"^\d+\.", plain_text(line)):
            seen_question = True
        if seen_question:
            question_lines.append(line)
        else:
            bank_lines.append(replace_blank_spans_with_spaces(line))
    filled = [replace_blank_spans_with_spaces(line) for line in bank_lines] + fill_lines_with_answers(question_lines, values)
    result: list[RichLine] = []
    for index, line in enumerate(filled):
        text = plain_text(line)
        if not text:
            continue
        result.append(RichLine(line, "phrase_bank" if not re.match(r"^\d+\.", text) else "body"))
    return result


def sentence_lines(lines: list[str], answer_lines: list[str]) -> list[RichLine]:
    values = parse_parenthetical_answers(answer_lines)
    filled = fill_numbered_blank_lines(lines, values)
    result: list[RichLine] = []
    for line in filled:
        text = plain_text(line)
        if not text:
            continue
        if text in {"句式分析", "句子仿写"}:
            result.append(RichLine(text, "exercise_head"))
        elif text.startswith("（") or text.startswith("("):
            result.append(RichLine(line, "body"))
        else:
            result.append(RichLine(line, "body"))
    return result


def writing_lines(lines: list[str], answer_lines: list[str]) -> list[RichLine]:
    numbered = parse_numbered_answers(answer_lines)
    values = [numbered[i] for i in sorted(numbered)]
    filled = fill_lines_with_answers(lines, values)
    result: list[RichLine] = []
    for line in filled:
        text = plain_text(line)
        if text:
            result.append(RichLine(line, "writing"))
    return result


def build_section_lines(
    title: str,
    body_lines: list[str],
    answer_lines: list[str],
    vocab_terms: list[str] | None = None,
    booklet_entries: list[str] | None = None,
) -> list[RichLine]:
    if title == "目标词汇":
        return target_vocab_lines(body_lines, answer_lines, booklet_entries)
    if title == "语法填空":
        return grammar_lines(body_lines, answer_lines, vocab_terms or [])
    if title == "核对译文":
        return translation_lines(body_lines)
    if title == "重点词汇拓展":
        return expansion_lines(body_lines, answer_lines)
    if title == "词块填空":
        return phrase_lines(body_lines, answer_lines)
    if title == "句型突破":
        return sentence_lines(body_lines, answer_lines)
    if title == "写作演练":
        return writing_lines(body_lines, answer_lines)
    return [RichLine(replace_blank_spans(line), "body") for line in body_lines if plain_text(line)]


def estimated_line_units(line: RichLine) -> float:
    text = plain_text(line.text)
    units = 0.0
    for ch in text:
        if ch in "\r\n":
            continue
        if ord(ch) > 127:
            units += 1.0
        elif ch.isspace():
            units += 0.35
        else:
            units += 0.55
    return max(1.0, units)


def line_cost(line: RichLine) -> float:
    role = line.role
    if role in {"article", "writing"}:
        width = 34.0
    elif role == "body":
        width = 38.0
    else:
        width = 46.0
    base = estimated_line_units(line) / width
    if role in {"article_title", "zh_title", "exercise_head"}:
        base += 0.4
    if role == "phrase_bank":
        base += 0.3
    return max(0.8, base)


def chunk_lines(lines: list[RichLine], title: str) -> list[list[RichLine]]:
    if not lines:
        return [[]]
    if title == "句型突破":
        chunks: list[list[RichLine]] = []
        current: list[RichLine] = []
        for line in lines:
            starts_new_pattern = (
                current
                and 'class="uline"' in line.text
                and "answer-fill" not in line.text
            )
            if starts_new_pattern:
                chunks.append(current)
                current = []
            current.append(line)
        if current:
            chunks.append(current)
        return chunks
    max_cost = {
        "目标词汇": 12.5,
        "语法填空": 12.0,
        "核对译文": 30.0,
        "重点词汇拓展": 8.0,
        "词块填空": 6.8,
        "句型突破": 16.0,
        "写作演练": 28.0,
    }.get(title, 10.5)
    chunks: list[list[RichLine]] = []
    current: list[RichLine] = []
    cost = 0.0
    for line in lines:
        item_cost = line_cost(line)
        if (
            title == "重点词汇拓展"
            and "answer-fill" in line.text
            and current
            and cost + item_cost > max_cost
        ):
            previous = current.pop()
            previous_cost = line_cost(previous)
            if current:
                chunks.append(current)
            current = [previous, line]
            cost = previous_cost + item_cost
            continue
        if current and cost + item_cost > max_cost:
            chunks.append(current)
            current = []
            cost = 0.0
        current.append(line)
        cost += item_cost
    if current:
        chunks.append(current)
    return chunks


def parse_inline_runs(text: str) -> list[HtmlRun]:
    text = re.sub(r'<span\s+class="blank"([^>]*)></span>', lambda m: blank_underscores(m.group(1) or ""), text)
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    text = italicize_pos_tokens(text)
    parser = SimpleInlineParser()
    try:
        parser.feed(text)
        parser.close()
    except Exception:
        return [HtmlRun(re.sub(r"<[^>]+>", "", text))]
    return parser.runs or [HtmlRun(re.sub(r"<[^>]+>", "", text))]


def italicize_pos_tokens(text: str) -> str:
    """Italicize part-of-speech abbreviations without touching HTML tags."""
    pieces = re.split(r"(<[^>]+>)", text)
    out: list[str] = []
    in_italic = False
    for piece in pieces:
        if piece.startswith("<") and piece.endswith(">"):
            tag = piece.lower()
            if re.match(r"<\s*(i|em)\b", tag):
                in_italic = True
            elif re.match(r"<\s*/\s*(i|em)\s*>", tag):
                in_italic = False
            out.append(piece)
            continue
        out.append(piece if in_italic else POS_TOKEN_RE.sub(r"<i>\1</i>", piece))
    return "".join(out)


def role_style(role: str) -> tuple[str, float, str, bool]:
    if role == "cover_title":
        return CJK_FONT, 54, ORANGE, True
    if role == "cover_subtitle_cn":
        return CJK_FONT, 28, WHITE, True
    if role == "cover_subtitle_en":
        return LATIN_FONT, 28, WHITE, True
    if role == "directory":
        return CJK_FONT, 32, "6E747A", True
    if role == "section_title":
        return CJK_FONT, 54, "6E747A", True
    if role == "top_header":
        return CJK_FONT, 32, WHITE, False
    if role == "article_title":
        return LATIN_FONT, 20, BLACK, True
    if role == "article":
        return LATIN_FONT, 20, BLACK, False
    if role == "zh_title":
        return CJK_FONT, 20, BLACK, True
    if role == "translation":
        return CJK_FONT, 20, BLACK, False
    if role == "exercise_head":
        return CJK_FONT, 20, BLACK, True
    if role == "phrase_bank":
        return LATIN_FONT, 20, BLACK, False
    if role == "writing":
        return LATIN_FONT, 20, BLACK, False
    if role == "vocab":
        return LATIN_FONT, 20, BLACK, True
    return LATIN_FONT, 20, BLACK, False


def append_run(
    paragraph: ET.Element,
    text: str,
    *,
    font: str,
    size: float,
    color: str,
    bold: bool,
    underline: bool = False,
    italic: bool = False,
    alpha: int | None = None,
) -> None:
    if not text:
        return
    r = ET.SubElement(paragraph, wenyan.qn(wenyan.A_NS, "r"))
    attrs = {"lang": "zh-CN", "sz": str(int(size * 100))}
    if bold:
        attrs["b"] = "1"
    if italic:
        attrs["i"] = "1"
    if underline:
        attrs["u"] = "sng"
    r_pr = ET.SubElement(r, wenyan.qn(wenyan.A_NS, "rPr"), attrs)
    solid = ET.SubElement(r_pr, wenyan.qn(wenyan.A_NS, "solidFill"))
    srgb = ET.SubElement(solid, wenyan.qn(wenyan.A_NS, "srgbClr"), {"val": color})
    if alpha is not None:
        ET.SubElement(srgb, wenyan.qn(wenyan.A_NS, "alpha"), {"val": str(alpha)})
    if underline:
        u_fill = ET.SubElement(r_pr, wenyan.qn(wenyan.A_NS, "uFill"))
        u_solid = ET.SubElement(u_fill, wenyan.qn(wenyan.A_NS, "solidFill"))
        ET.SubElement(u_solid, wenyan.qn(wenyan.A_NS, "srgbClr"), {"val": "808080" if color == ORANGE else BLACK})
    ET.SubElement(r_pr, wenyan.qn(wenyan.A_NS, "latin"), {"typeface": LATIN_FONT})
    ET.SubElement(r_pr, wenyan.qn(wenyan.A_NS, "ea"), {"typeface": CJK_FONT})
    ET.SubElement(r_pr, wenyan.qn(wenyan.A_NS, "cs"), {"typeface": LATIN_FONT})
    t = ET.SubElement(r, wenyan.qn(wenyan.A_NS, "t"))
    if text[:1].isspace() or text[-1:].isspace() or "  " in text:
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text


def append_textbox(
    root: ET.Element,
    *,
    x: float,
    y: float,
    w: float,
    h: float,
    lines: list[RichLine],
    name: str,
    align: str = "l",
    line_spacing_pt: float | None = CONTENT_LINE_SPACING_PT,
    answer_mode: str = "normal",
    target_answer_index: int | None = None,
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
    ET.SubElement(
        tx_body,
        wenyan.qn(wenyan.A_NS, "bodyPr"),
        {"wrap": "square", "rtlCol": "0", "lIns": "0", "tIns": "0", "rIns": "0", "bIns": "0"},
    )
    ET.SubElement(tx_body, wenyan.qn(wenyan.A_NS, "lstStyle"))
    answer_index = 0
    for line in lines or [RichLine("")]:
        font, size, color, bold = role_style(line.role)
        p = ET.SubElement(tx_body, wenyan.qn(wenyan.A_NS, "p"))
        p_attrs = {"algn": align, "marL": "0", "indent": "0"}
        if line.role == "article":
            p_attrs["indent"] = str(wenyan.inch(GRAMMAR_FIRST_LINE_INDENT_IN))
        p_pr = ET.SubElement(p, wenyan.qn(wenyan.A_NS, "pPr"), p_attrs)
        if line_spacing_pt:
            ln_spc = ET.SubElement(p_pr, wenyan.qn(wenyan.A_NS, "lnSpc"))
            ET.SubElement(ln_spc, wenyan.qn(wenyan.A_NS, "spcPts"), {"val": str(int(line_spacing_pt * 100))})
        for run in parse_inline_runs(line.text):
            run_color = ORANGE if run.answer else color
            run_bold = True if run.answer else (bold or run.bold)
            run_text = run.text
            run_underline = run.underline
            run_alpha: int | None = None
            if answer_mode == "base" and run.answer:
                run_color = BLACK
                run_bold = True
                run_underline = run.underline
                run_alpha = 0
            elif answer_mode == "overlay":
                current_answer_index = answer_index if run.answer else None
                if run.answer and current_answer_index == target_answer_index:
                    run_color = ORANGE
                    run_bold = True
                    run_underline = False
                else:
                    run_alpha = 0
                    run_underline = False
            if run.answer:
                answer_index += 1
            append_run(
                p,
                run_text,
                font=font,
                size=size,
                color=run_color,
                bold=run_bold,
                underline=run_underline,
                italic=run.italic,
                alpha=run_alpha,
            )
        ET.SubElement(p, wenyan.qn(wenyan.A_NS, "endParaRPr"), {"lang": "zh-CN", "sz": str(int(size * 100))})
    wenyan.append_to_sp_tree(root, sp)
    return shape_id


def append_link_textbox(
    root: ET.Element,
    rels_root: ET.Element,
    *,
    x: float,
    y: float,
    w: float,
    h: float,
    text: str,
    target_slide: int,
    name: str,
) -> int:
    shape_id = wenyan.next_shape_id(root)
    rel_id = wenyan.add_slide_jump_relationship(rels_root, target_slide)
    sp = ET.Element(wenyan.qn(wenyan.P_NS, "sp"))
    nv_sp_pr = ET.SubElement(sp, wenyan.qn(wenyan.P_NS, "nvSpPr"))
    c_nv_pr = ET.SubElement(nv_sp_pr, wenyan.qn(wenyan.P_NS, "cNvPr"), {"id": str(shape_id), "name": f"{name} {shape_id}"})
    ET.SubElement(c_nv_pr, wenyan.qn(wenyan.A_NS, "hlinkClick"), {wenyan.qn(wenyan.R_NS, "id"): rel_id})
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
    ET.SubElement(tx_body, wenyan.qn(wenyan.A_NS, "bodyPr"), {"wrap": "none", "rtlCol": "0"})
    ET.SubElement(tx_body, wenyan.qn(wenyan.A_NS, "lstStyle"))
    p = ET.SubElement(tx_body, wenyan.qn(wenyan.A_NS, "p"))
    ET.SubElement(p, wenyan.qn(wenyan.A_NS, "pPr"), {"algn": "l"})
    r = ET.SubElement(p, wenyan.qn(wenyan.A_NS, "r"))
    r_pr = ET.SubElement(r, wenyan.qn(wenyan.A_NS, "rPr"), {"lang": "zh-CN", "sz": "3200", "b": "1"})
    solid = ET.SubElement(r_pr, wenyan.qn(wenyan.A_NS, "solidFill"))
    ET.SubElement(solid, wenyan.qn(wenyan.A_NS, "srgbClr"), {"val": "6E747A"})
    ET.SubElement(r_pr, wenyan.qn(wenyan.A_NS, "latin"), {"typeface": LATIN_FONT})
    ET.SubElement(r_pr, wenyan.qn(wenyan.A_NS, "ea"), {"typeface": CJK_FONT})
    ET.SubElement(r_pr, wenyan.qn(wenyan.A_NS, "cs"), {"typeface": LATIN_FONT})
    t = ET.SubElement(r, wenyan.qn(wenyan.A_NS, "t"))
    t.text = text
    ET.SubElement(p, wenyan.qn(wenyan.A_NS, "endParaRPr"), {"lang": "zh-CN", "sz": "3200"})
    wenyan.append_to_sp_tree(root, sp)
    return shape_id


def count_answer_runs(lines: list[RichLine]) -> int:
    return sum(1 for line in lines for run in parse_inline_runs(line.text) if run.answer)


def append_answer_aware_textbox(
    root: ET.Element,
    *,
    x: float,
    y: float,
    w: float,
    h: float,
    lines: list[RichLine],
    name: str,
    align: str = "l",
    line_spacing_pt: float | None = CONTENT_LINE_SPACING_PT,
) -> None:
    answer_count = count_answer_runs(lines)
    if answer_count <= 0:
        append_textbox(
            root,
            x=x,
            y=y,
            w=w,
            h=h,
            lines=lines,
            name=name,
            align=align,
            line_spacing_pt=line_spacing_pt,
        )
        return
    append_textbox(
        root,
        x=x,
        y=y,
        w=w,
        h=h,
        lines=lines,
        name=f"{name} base",
        align=align,
        line_spacing_pt=line_spacing_pt,
        answer_mode="base",
    )
    for index in range(answer_count):
        append_textbox(
            root,
            x=x,
            y=y,
            w=w,
            h=h,
            lines=lines,
            name=f"{name} answer {index + 1}",
            align=align,
            line_spacing_pt=line_spacing_pt,
            answer_mode="overlay",
            target_answer_index=index,
        )


def append_solid_line(
    root: ET.Element,
    *,
    x: float,
    y: float,
    w: float,
    color: str = BLACK,
    name: str = "english answer line",
) -> int:
    shape_id = wenyan.next_shape_id(root)
    cxn = ET.Element(wenyan.qn(wenyan.P_NS, "cxnSp"))
    nv_pr = ET.SubElement(cxn, wenyan.qn(wenyan.P_NS, "nvCxnSpPr"))
    ET.SubElement(nv_pr, wenyan.qn(wenyan.P_NS, "cNvPr"), {"id": str(shape_id), "name": f"{name} {shape_id}"})
    ET.SubElement(nv_pr, wenyan.qn(wenyan.P_NS, "cNvCxnSpPr"))
    ET.SubElement(nv_pr, wenyan.qn(wenyan.P_NS, "nvPr"))

    sp_pr = ET.SubElement(cxn, wenyan.qn(wenyan.P_NS, "spPr"))
    xfrm = ET.SubElement(sp_pr, wenyan.qn(wenyan.A_NS, "xfrm"))
    ET.SubElement(xfrm, wenyan.qn(wenyan.A_NS, "off"), {"x": str(wenyan.inch(x)), "y": str(wenyan.inch(y))})
    ET.SubElement(xfrm, wenyan.qn(wenyan.A_NS, "ext"), {"cx": str(wenyan.inch(w)), "cy": "0"})
    geom = ET.SubElement(sp_pr, wenyan.qn(wenyan.A_NS, "prstGeom"), {"prst": "line"})
    ET.SubElement(geom, wenyan.qn(wenyan.A_NS, "avLst"))
    ln = ET.SubElement(sp_pr, wenyan.qn(wenyan.A_NS, "ln"), {"w": "12700"})
    solid = ET.SubElement(ln, wenyan.qn(wenyan.A_NS, "solidFill"))
    ET.SubElement(solid, wenyan.qn(wenyan.A_NS, "srgbClr"), {"val": color})
    wenyan.append_to_sp_tree(root, cxn)
    return shape_id


def split_vocab_answer(line: RichLine) -> tuple[str, str | None]:
    match = re.search(r'<span\s+class="answer-fill"[^>]*>(.*?)</span>', line.text)
    if not match:
        return line.text, None
    prompt = line.text[: match.start()].strip()
    answer = re.sub(r"<[^>]+>", "", match.group(1))
    return prompt, html.unescape(answer).strip()


def answer_line_width(answer: str | None) -> float:
    if not answer:
        return 1.45
    units = 0.0
    for ch in answer:
        if ord(ch) > 127:
            units += 1.0
        elif ch.isspace():
            units += 0.35
        else:
            units += 0.55
    return max(1.55, min(9.4, units * 0.30 + 0.55))


def prompt_width(prompt: str) -> float:
    text = plain_text(prompt)
    units = 0.0
    for ch in text:
        if ord(ch) > 127:
            units += 1.0
        elif ch.isspace():
            units += 0.35
        else:
            units += 0.55
    return max(1.55, min(3.65, units * 0.24 + 0.3))


def add_target_vocab_content(root: ET.Element, lines: list[RichLine]) -> None:
    x = 0.82
    y = 1.88
    row_gap = 0.43
    right_edge = 12.62
    for index, line in enumerate(lines):
        row_y = y + index * row_gap
        prompt, answer = split_vocab_answer(line)
        q_w = prompt_width(prompt)
        append_textbox(
            root,
            x=x,
            y=row_y,
            w=q_w + 0.12,
            h=0.38,
            lines=[RichLine(prompt, "vocab")],
            name="english target vocab prompt",
            line_spacing_pt=None,
        )
        line_x = x + q_w + 0.08
        line_w = min(answer_line_width(answer), max(1.2, right_edge - line_x))
        append_solid_line(root, x=line_x, y=row_y + 0.32, w=line_w, name="english target vocab underline")
        if answer:
            append_textbox(
                root,
                x=line_x,
                y=row_y - 0.02,
                w=line_w,
                h=0.38,
                lines=[RichLine(f'<span class="answer-fill">{html.escape(answer, quote=False)}</span>', "vocab")],
                name=f"english target vocab answer {index + 1}",
                align="ctr",
                line_spacing_pt=None,
            )


def append_header(root: ET.Element, title: str) -> None:
    append_textbox(
        root,
        x=1.13,
        y=0.76,
        w=5.6,
        h=0.77,
        lines=[RichLine(title, "top_header")],
        name="english exercise header",
        line_spacing_pt=None,
    )


def add_cover(root: ET.Element, passage: Passage) -> None:
    for sp in wenyan.text_shapes(root):
        if wenyan.get_shape_text(sp).strip() == "人与自我      Passage 1":
            wenyan.set_shape_text(sp, "")
    append_textbox(
        root,
        x=0.28,
        y=2.54,
        w=8.57,
        h=1.01,
        lines=[RichLine(BOOK_TITLE, "cover_title")],
        name="english cover title",
        line_spacing_pt=None,
    )
    append_textbox(
        root,
        x=7.97,
        y=6.71,
        w=1.7,
        h=0.57,
        lines=[RichLine(passage.theme, "cover_subtitle_cn")],
        name="english cover theme",
        line_spacing_pt=None,
    )
    append_textbox(
        root,
        x=9.75,
        y=6.71,
        w=2.9,
        h=0.57,
        lines=[RichLine(f"Passage {passage.num}", "cover_subtitle_en")],
        name="english cover passage",
        line_spacing_pt=None,
    )


def add_directory(root: ET.Element, rels_root: ET.Element, title_slide_numbers: dict[str, int]) -> None:
    positions = {
        "目标词汇": (2.55, 2.05),
        "语法填空": (7.15, 2.05),
        "核对译文": (2.55, 3.05),
        "重点词汇拓展": (7.15, 3.05),
        "词块填空": (2.55, 4.05),
        "句型突破": (7.15, 4.05),
        "写作演练": (2.55, 5.05),
    }
    for title in EXERCISE_TITLES:
        x, y = positions[title]
        append_link_textbox(
            root,
            rels_root,
            x=x,
            y=y,
            w=3.45 if title == "重点词汇拓展" else 2.8,
            h=0.48,
            text=f"{EXERCISE_NUMBERS[title]}  {title}",
            target_slide=title_slide_numbers[title],
            name=f"english directory {title}",
        )


def add_section_title(root: ET.Element, title: str) -> None:
    label = f"{EXERCISE_NUMBERS[title]}  {title}"
    append_textbox(
        root,
        x=3.52,
        y=3.25,
        w=5.53,
        h=1.01,
        lines=[RichLine(label, "section_title")],
        name="english section title",
        align="ctr",
        line_spacing_pt=None,
    )


def add_content(root: ET.Element, title: str, lines: list[RichLine]) -> None:
    header = f"{EXERCISE_NUMBERS[title]} {title}"
    append_header(root, header)
    y = 2.09
    h = CONTENT_BOTTOM_Y - y
    if title == "目标词汇":
        add_target_vocab_content(root, lines)
        return
    if title == "词块填空" and lines and lines[0].role == "phrase_bank":
        bank: list[RichLine] = []
        rest: list[RichLine] = []
        for line in lines:
            (bank if line.role == "phrase_bank" and not rest else rest).append(line)
        if bank:
            wenyan.append_rect(root, x=0.35, y=1.85, w=12.55, h=1.05, fill_color=None, line_color="808080", name="phrase bank border")
            append_textbox(root, x=2.8, y=2.05, w=7.4, h=0.85, lines=bank, name="english phrase bank", align="ctr")
            append_answer_aware_textbox(root, x=0.38, y=3.05, w=12.0, h=CONTENT_BOTTOM_Y - 3.05, lines=rest, name="english phrase content")
            return
    if title == "语法填空":
        if lines and lines[0].role == "article_title":
            append_textbox(root, x=0.46, y=2.05, w=12.87, h=0.35, lines=[lines[0]], name="english grammar title", align="ctr")
            append_answer_aware_textbox(root, x=0.46, y=2.43, w=12.87, h=CONTENT_BOTTOM_Y - 2.43, lines=lines[1:], name="english grammar body")
            return
        append_answer_aware_textbox(root, x=0.46, y=2.09, w=12.87, h=CONTENT_BOTTOM_Y - 2.09, lines=lines, name="english grammar body")
        return
    if title == "写作演练":
        y = 2.0
        h = CONTENT_BOTTOM_Y - y
    append_answer_aware_textbox(root, x=0.42, y=y, w=12.35, h=h, lines=lines, name=f"english {title} content")


def template_slide_count(template: Path) -> int:
    with zipfile.ZipFile(template, "r") as zf:
        return len([name for name in zf.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)])


def write_deck(
    template: Path,
    output: Path,
    slide_specs: list[tuple[int, Callable[[ET.Element, ET.Element], None]]],
) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    generated: list[tuple[int, bytes, bytes]] = []
    with zipfile.ZipFile(template, "r") as zin:
        for src, edit in slide_specs:
            xml, rels = wenyan.make_slide_parts(zin, src, edit)
            generated.append((src, xml, rels))
        slide_total = len(generated)
        content_types = wenyan.update_content_types(wenyan.parse_xml(zin.read("[Content_Types].xml")), slide_total)
        presentation_xml = wenyan.update_presentation_xml(wenyan.parse_xml(zin.read("ppt/presentation.xml")), slide_total)
        presentation_rels = wenyan.update_presentation_rels(wenyan.parse_xml(zin.read("ppt/_rels/presentation.xml.rels")), slide_total)
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            skip = {"[Content_Types].xml", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"}
            for item in zin.infolist():
                name = item.filename
                if item.is_dir() or name in skip:
                    continue
                if re.match(r"ppt/slides/slide\d+\.xml$", name):
                    continue
                if re.match(r"ppt/slides/_rels/slide\d+\.xml\.rels$", name):
                    continue
                zout.writestr(item, zin.read(name))
            zout.writestr("[Content_Types].xml", content_types)
            zout.writestr("ppt/presentation.xml", presentation_xml)
            zout.writestr("ppt/_rels/presentation.xml.rels", presentation_rels)
            for index, (_src, xml, rels) in enumerate(generated, 1):
                zout.writestr(f"ppt/slides/slide{index}.xml", xml)
                zout.writestr(f"ppt/slides/_rels/slide{index}.xml.rels", rels)


def rebuild_answer_animations_with_powerpoint(pptx: Path) -> int:
    try:
        import win32com.client
    except Exception as exc:
        print(f"提示：未能导入 win32com，答案动画未重建：{exc}", file=sys.stderr)
        return 0

    pp = None
    pres = None
    rebuilt = 0
    try:
        pp = win32com.client.DispatchEx("PowerPoint.Application")
        pres = pp.Presentations.Open(str(pptx), WithWindow=False, ReadOnly=False)
        for slide in pres.Slides:
            rebuilt += wenyan.add_powerpoint_standard_animations(slide)
        if rebuilt:
            pres.Save()
    except Exception as exc:
        print(f"提示：PowerPoint 答案动画重建失败：{exc}", file=sys.stderr)
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
    return rebuilt


def find_md_files(md_run: Path) -> tuple[Path, Path, Path | None]:
    body = sorted(md_run.glob("正文_*.md"))
    answer = sorted(md_run.glob("答案_*.md"))
    booklet = sorted(md_run.glob("册子_*.md")) or sorted(md_run.glob("*册子*.md"))
    if not body or not answer:
        raise FileNotFoundError(f"未找到正文/答案 Markdown：{md_run}")
    return body[0], answer[0], booklet[0] if booklet else None


def generate_english_passage_pptx(
    *,
    md_run: Path,
    template: Path,
    output_dir: Path,
    passage_num: int = 1,
    progress: Callable[[str], None] | None = None,
) -> Path:
    def emit(message: str) -> None:
        if progress:
            progress(message)
        else:
            print(message)

    body_md, answer_md, booklet_md = find_md_files(md_run)
    passages = parse_body_passages(body_md)
    answers = parse_answer_passages(answer_md)
    booklet_vocab = parse_booklet_vocab_passages(booklet_md)
    if not passages:
        raise ValueError(f"未解析到 passage：{body_md}")
    passage = next((item for item in passages if item.num == passage_num), passages[0])
    passage_answers = answers.get(passage.num, PassageAnswers({}))
    emit(f"正文 MD：{body_md}")
    emit(f"答案 MD：{answer_md}")
    if booklet_md:
        emit(f"册子 MD：{booklet_md}")
    else:
        emit("册子 MD：未找到，目标词汇答案将使用旧兜底逻辑。")
    emit(f"解析到 passage 数：{len(passages)}，本次生成 Passage {passage.num}")

    built_sections: list[BuiltSection] = []
    next_slide_number = 3
    vocab_terms = extract_vocab_terms(passage.sections.get("目标词汇", []))
    for title in EXERCISE_TITLES:
        body_lines = passage.sections.get(title, [])
        answer_key = EXERCISE_TO_ANSWER.get(title)
        answer_lines = passage_answers.exercises.get(answer_key, []) if answer_key is not None else []
        rich_lines = build_section_lines(
            title,
            body_lines,
            answer_lines,
            vocab_terms,
            booklet_vocab.get(passage.num),
        )
        chunks = chunk_lines(rich_lines, title)
        built_sections.append(BuiltSection(title, chunks, next_slide_number))
        next_slide_number += 1 + len(chunks)

    title_slide_numbers = {section.title: section.title_slide_number for section in built_sections}

    slide_specs: list[tuple[int, Callable[[ET.Element, ET.Element], None]]] = []
    slide_specs.append((1, lambda root, rels: add_cover(root, passage)))
    slide_specs.append((2, lambda root, rels: add_directory(root, rels, title_slide_numbers)))
    for section in built_sections:
        title = section.title
        slide_specs.append((3, lambda root, rels, title=title: add_section_title(root, title)))
        for chunk in section.chunks:
            slide_specs.append((4, lambda root, rels, title=title, chunk=chunk: add_content(root, title, chunk)))
    for src in (5, 6, 7):
        slide_specs.append((src, lambda root, rels: None))

    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / f"Passage {passage.num}.pptx"
    write_deck(template, output, slide_specs)
    rebuilt = rebuild_answer_animations_with_powerpoint(output)
    if rebuilt:
        emit(f"答案动画：已重建 {rebuilt} 个动画对象")
    emit(f"已生成：{output}")
    return output


def generate_english_book_pptx(
    *,
    md_run: Path,
    template: Path,
    output_dir: Path,
    passage_nums: Iterable[int] | None = None,
    progress: Callable[[str], None] | None = None,
) -> list[Path]:
    """Generate one PPTX per passage for the English vocabulary book."""

    def emit(message: str) -> None:
        if progress:
            progress(message)
        else:
            print(message)

    body_md, _answer_md, _booklet_md = find_md_files(md_run)
    passages = parse_body_passages(body_md)
    if not passages:
        raise ValueError(f"No passage parsed from {body_md}")

    wanted = set(passage_nums) if passage_nums else None
    selected = [passage for passage in passages if wanted is None or passage.num in wanted]
    if not selected:
        raise ValueError(f"No requested passage found: {sorted(wanted) if wanted else []}")

    emit(f"English vocabulary book: generating {len(selected)} Passage PPTX files.")
    outputs: list[Path] = []
    for index, passage in enumerate(selected, 1):
        emit(f"[{index}/{len(selected)}] Generate Passage {passage.num}")
        outputs.append(
            generate_english_passage_pptx(
                md_run=md_run,
                template=template,
                output_dir=output_dir,
                passage_num=passage.num,
                progress=progress,
            )
        )
    emit(f"English PPT generation complete: {output_dir}")
    return outputs


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="测试：英语词汇语境记忆 md + 模板生成单个 passage PPTX。")
    parser.add_argument("--md-run", default=str(DEFAULT_MD_RUN), help="英语 md run 目录。")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="英语 PPTX 模板路径。")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="测试 PPTX 输出目录。")
    parser.add_argument("--passage", type=int, default=1, help="生成第几个 passage，默认 1。")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    output_dir = Path(args.output_dir)
    if output_dir == DEFAULT_OUTPUT_DIR:
        output_dir = output_dir.with_name(f"test_english_ppt_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    generate_english_passage_pptx(
        md_run=Path(args.md_run),
        template=Path(args.template),
        output_dir=output_dir,
        passage_num=args.passage,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

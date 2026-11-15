#!/usr/bin/env python3
"""《高考语文大一轮复习新方案》测试生成器。

本适配器只负责新书的专题化结构：

* 正文按“板块必备 / 专题突破 / 专题一 / 专题二”切分；
* 答案按“专题 n -> 突破 n”建立局部答案映射，避免不同突破中重复的
  1、2、3 题互相串答案；
* 具体幻灯片 XML、答案动画、文本框裁剪、TIF 放置和模板克隆复用
  tests/test_md_to_ppt.py 与文言文阅读类测试程序。

当前测试目标：1阅读Ⅰ 信息类文本阅读（专题一—专题二）。
"""

from __future__ import annotations

import argparse
import json
import math
import posixpath
import re
import sys
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
import test_wenyan_yuedu_jiafen_md_to_ppt as yuedu  # noqa: E402
from fbd_to_md import convert_fbd_to_markdown, read_fbd, tif_markdown_tag  # noqa: E402
from generation_state import isolated_generation  # noqa: E402


DEFAULT_SOURCE_DIR = PROJECT_DIR / "高考语文大一轮复习新方案_源文件"
DEFAULT_BODY_FBD = DEFAULT_SOURCE_DIR / "一轮新方案.fbd"
DEFAULT_ANSWER_FBD = DEFAULT_SOURCE_DIR / "一轮新方案答案.fbd"
DEFAULT_TEMPLATE = PROJECT_DIR / "大一轮复习新方案模板.pptx"
DEFAULT_MD_ROOT = PROJECT_DIR / "out_md" / "高考语文大一轮复习新方案"
DEFAULT_OUTPUT_DIR = PROJECT_DIR / "大一轮复习pptx" / "_生成"

BOOK_TITLE = "高考语文大一轮复习新方案"
TARGET_VOLUME = "阅读Ⅰ 信息类文本阅读"
READING_II = "阅读Ⅱ 文学类文本阅读"
MAIN_MARKER = sanbu.MAIN_MARKER
# The shared renderer uses a compact estimate for long-answer lines.  In this
# book those lines include a long editable underline and usually occupy two
# visual rows, so keep a little more vertical headroom on every content slide.
YILUN_MAX_PAGE_UNITS = 10
# 判断题的题干、原文和“设错角度及理由”答案是一个连续版块。
# 这一版块的文字量较大，单独提高容量，避免在页面仍有大块空白时提前换页。
YILUN_JUDGEMENT_MAX_PAGE_UNITS = 26
YILUN_LONG_ANSWER_PAGE_COST = 2
YILUN_JUDGEMENT_FIRST_ANSWER_UNITS = 34.0
YILUN_JUDGEMENT_CONTINUED_ANSWER_UNITS = 42.0
YILUN_LARGE_TIF_SPACER_LINES = 8
YILUN_LARGE_TIF_MAX_HEIGHT = 3.18
YILUN_LARGE_TIF_PORTRAIT_MAX_HEIGHT = 3.55
YILUN_TIF_SCALE = 0.5
NBSP = "\u00a0"

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

YILUN_TABLE_SPECS = {
    "内涵特点": {
        "headers": ["内涵", "特点"],
        "mode": "first_rest",
        "stop": r"^2[\.．、]?非连续性文本材料间|^材料关系",
    },
    "材料关系具体特征材料话题举例材料一材料二材料三": {
        "headers": ["材料关系", "具体特征", "材料话题举例", "材料一", "材料二", "材料三"],
        "row_re": r"^(并列|互补|对立|总分/分总|综合)",
        "stop": r"^说明[:：]|^\(?二\)?连续性文本",
    },
    "类别内涵特点": {
        "headers": ["类别", "内涵", "特点"],
        "row_re": r"^(政论文|学术论文|文艺评论|时评|书评|新闻|科普文|报告|访谈|传记)",
        "stop": r"^2[\.．、]?实用类文本|^三$|^高考解读|^选文概览",
    },
    "选文概览命题概览": {
        "headers": ["卷别", "选材/文本类型", "命题概览"],
        "skip_next_header": "卷别选材文本类型T1T2T3T4T5",
        "row_re": r"^20\d{2}·|^2025·八省",
        "stop": r"^试题研究$",
    },
    "试题研究": {
        "headers": ["项目", "内容"],
        "row_re": r"^(选材特点|命题特点|命题趋势|备考策略)",
        "stop": r"^[一二三四五六七八九十]+、|^专题突破$",
    },
    "卷别提问方式能力点": {
        "headers": ["卷别", "提问方式", "能力点"],
        "row_re": r"^20\d{2}·|^2025·八省|^2024·全国|^2023·北京",
        "stop": r"^\d+$|^特别提醒[:：]",
    },
    "转述方式特点解题要点": {
        "headers": ["转述方式", "特点", "解题要点"],
        "row_re": r"^(概括式|具体式|同义变换|关联式)",
        "stop": r"^[一二三四五六七八九十]+、|^判断下列",
    },
    "类型角度阐释识别技巧": {
        "headers": ["类型/角度", "阐释与识别技巧"],
        "row_re": r"^(信息曲解|混淆类|曲解文意|偷换概念/|张冠李戴|混淆是非|混淆时态|信息转述|偏颇类|以偏概全/|范围不清|绝对武断|无中生有/|于文无据|逻辑错乱)",
        "stop": r"^判断下列|^\d+\.",
    },
    "命题角度解题方法": {
        "headers": ["命题角度", "解题方法"],
        "row_re": r"^(推断前提是否有据|推断过程是否合乎逻辑|推断结论是否合理)",
        "stop": r"^[一二三四五六七八九十]+、|^\d+\.",
    },
    "形式特点教材示例": {
        "headers": ["形式", "特点/教材示例"],
        "row_re": r"^(总分式|并列式|递进式|对照式)",
        "stop": r"^递进式\(层进式\)|^[一二三四五六七八九十]+、",
    },
    "角度具体阐释": {
        "headers": ["角度", "具体阐释"],
        "row_re": r"^(论点提出的|位置或方法|论据特点|论证结构|论证方式|论证方法|论证语言)",
        "stop": r"^[一二三四五六七八九十]+、|^二、",
    },
    "论证方法特点作用教材示例": {
        "headers": ["论证方法", "特点/作用/教材示例"],
        "row_re": r"^(举例论证|对比论证|比喻论证|引用论证|因果论证|假设论证|类比论证|归谬论证)",
        "stop": r"^[一二三四五六七八九十]+、|^三、|^\d+\.",
    },
    "方法具体阐释": {
        "headers": ["方法", "具体阐释"],
        "row_re": r"^(提取法|合并法|转换法|比较法|筛选法|整合法)",
        "stop": r"^[一二三四五六七八九十]+、|^\d+\.",
    },
    "特点具体阐释": {
        "headers": ["特点", "具体阐释"],
        "row_re": r"^(题干表|述复杂|设题情|境化|答题区域|跨文本|题型丰富)",
        "stop": r"^\(\d+\)|^[一二三四五六七八九十]+、",
    },
    "卷别选材话题及来源材料类别提问方式能力点": {
        "headers": ["卷别", "选材话题及来源/材料类别/提问方式/能力点"],
        "row_re": r"^20\d{2}·|^2025·八省",
        "stop": r"^\d+$|^特别提醒[:：]",
    },
}

YILUN_TABLE_HEADER_KEYS = set(YILUN_TABLE_SPECS)
# FBD 中的表格标记在转换为普通 Markdown 时会被清掉。测试生成器在
# 有源 FBD 时额外读取 BG/BH 结构，把未手工登记的表格恢复为原生 PPT
# 表格；键相同的表格保留为候选列表，以支持同一书中重复表头。
YILUN_DYNAMIC_TABLE_SPECS: dict[str, list[dict[str, object]]] = {}

YILUN_ABILITY_TERMS = (
    "信息的理解",
    "信息的评价",
    "信息的推断",
    "辨析论点和论据",
    "分析论证方法及论证效果",
    "分析论证方法",
    "综合分析论证特点",
    "分析论证语言",
    "比较材料异同",
    "概括内容要点",
    "理解概念含义、概括内容要点",
    "理解概念含义",
    "比较材料异同",
    "迁移运用观点",
    "分析评价观点",
    "准确信息推断",
    "比较分析异同",
)

YILUN_BOARD_SPECS = [
    {
        "key": "阅读Ⅰ 信息类文本阅读",
        "start": "阅读Ⅰ 信息类文本阅读",
        "end": "阅读Ⅱ 文学类文本阅读",
        "volume": "阅读Ⅰ 信息类文本阅读",
        "strip_titles": ("阅读Ⅰ 信息类文本阅读",),
        "num_offset": 0,
    },
    {
        "key": "阅读Ⅱ 文学类文本阅读小说",
        "start": "板块一 小说",
        "end": "板块二 散文",
        "volume": "阅读Ⅱ 文学类文本阅读小说",
        "strip_titles": ("板块一 小说",),
        "num_offset": 0,
    },
    {
        "key": "阅读Ⅱ 文学类文本阅读散文",
        "start": "板块二 散文",
        "end": "阅读Ⅲ 文言文阅读",
        "volume": "阅读Ⅱ 文学类文本阅读散文",
        "strip_titles": ("板块二 散文",),
        "num_offset": 0,
    },
    {
        "key": "阅读Ⅲ 文言文阅读",
        "start": "阅读Ⅲ 文言文阅读",
        "end": "阅读Ⅳ 古代诗歌阅读",
        "volume": "阅读Ⅲ 文言文阅读",
        "strip_titles": ("阅读Ⅲ 文言文阅读",),
        "num_offset": 0,
    },
    {
        "key": "阅读Ⅳ 古代诗歌阅读",
        "start": "阅读Ⅳ 古代诗歌阅读",
        "end": "名篇名句默写",
        "volume": "阅读Ⅳ 古代诗歌阅读",
        "strip_titles": ("阅读Ⅳ 古代诗歌阅读",),
        "num_offset": 0,
    },
    {
        "key": "阅读Ⅳ 古代诗歌阅读+默写",
        "start": "名篇名句默写",
        "end": "语言文字运用",
        "volume": "阅读Ⅳ 古代诗歌阅读+默写",
        "strip_titles": ("名篇名句默写",),
        "num_offset": 3,
    },
    {
        "key": "语言文字运用",
        "start": "语言文字运用",
        "end": "写作",
        "volume": "语言文字运用",
        "strip_titles": ("语言文字运用",),
        "num_offset": 0,
    },
    {
        "key": "写作",
        "start": "写作",
        "end": None,
        "volume": "写作",
        "strip_titles": ("写作",),
        "num_offset": 0,
    },
]

YILUN_DECK_SPECS = [
    {
        "filename": "1阅读Ⅰ 信息类文本阅读（专题一—专题二）.pptx",
        "board": "阅读Ⅰ 信息类文本阅读",
        "piece_start": 1,
        "piece_end": 3,
        "cover_start": 1,
        "cover_end": 2,
    },
    {
        "filename": "2阅读Ⅰ 信息类文本阅读（专题三—专题五）.pptx",
        "board": "阅读Ⅰ 信息类文本阅读",
        "piece_start": 4,
        "piece_end": 6,
        "cover_start": 3,
        "cover_end": 5,
    },
    {
        "filename": "3阅读Ⅱ 文学类文本阅读小说（专题一—专题二）.pptx",
        "board": "阅读Ⅱ 文学类文本阅读小说",
        "piece_start": 1,
        "piece_end": 3,
        "cover_start": 1,
        "cover_end": 2,
    },
    {
        "filename": "4阅读Ⅱ 文学类文本阅读小说（专题三—专题五）.pptx",
        "board": "阅读Ⅱ 文学类文本阅读小说",
        "piece_start": 4,
        "piece_end": 6,
        "cover_start": 3,
        "cover_end": 5,
    },
    {
        "filename": "5阅读Ⅱ 文学类文本阅读散文（专题一—专题三）.pptx",
        "board": "阅读Ⅱ 文学类文本阅读散文",
        "piece_start": 1,
        "piece_end": 4,
        "cover_start": 1,
        "cover_end": 3,
    },
    {
        "filename": "6阅读Ⅱ 文学类文本阅读散文（专题四—专题六）.pptx",
        "board": "阅读Ⅱ 文学类文本阅读散文",
        "piece_start": 5,
        "piece_end": 7,
        "cover_start": 4,
        "cover_end": 6,
    },
    {
        "filename": "7阅读Ⅲ 文言文阅读（专题一—专题三）.pptx",
        "board": "阅读Ⅲ 文言文阅读",
        "piece_start": 1,
        "piece_end": 4,
        "cover_start": 1,
        "cover_end": 3,
    },
    {
        "filename": "8阅读Ⅲ 文言文阅读（专题四—专题七）.pptx",
        "board": "阅读Ⅲ 文言文阅读",
        "piece_start": 5,
        "piece_end": 8,
        "cover_start": 4,
        "cover_end": 7,
    },
    {
        "filename": "9阅读Ⅳ 古代诗歌阅读（专题一—专题三）.pptx",
        "board": "阅读Ⅳ 古代诗歌阅读",
        "piece_start": 1,
        "piece_end": 4,
        "cover_start": 1,
        "cover_end": 3,
    },
    {
        "filename": "10阅读Ⅳ 古代诗歌阅读+默写（专题四—专题六，默写三个专题）.pptx",
        "board": "阅读Ⅳ 古代诗歌阅读+默写",
        "piece_start": 4,
        "piece_end": 6,
        "cover_start": 4,
        "cover_end": 6,
    },
    {
        "filename": "11语言文字运用（专题一—专题二）.pptx",
        "board": "语言文字运用",
        "piece_start": 1,
        "piece_end": 3,
        "cover_start": 1,
        "cover_end": 2,
    },
    {
        "filename": "12语言文字运用（专题三—专题七）.pptx",
        "board": "语言文字运用",
        "piece_start": 4,
        "piece_end": 8,
        "cover_start": 3,
        "cover_end": 7,
    },
    {
        "filename": "13语言文字运用（专题八—专题十）.pptx",
        "board": "语言文字运用",
        "piece_start": 9,
        "piece_end": 11,
        "cover_start": 8,
        "cover_end": 10,
    },
    {
        "filename": "14写作（专题一—专题二）.pptx",
        "board": "写作",
        "piece_start": 1,
        "piece_end": 3,
        "cover_start": 1,
        "cover_end": 2,
    },
    {
        "filename": "15写作（专题三—专题四）.pptx",
        "board": "写作",
        "piece_start": 4,
        "piece_end": 5,
        "cover_start": 3,
        "cover_end": 4,
    },
]


def plain(text: str) -> str:
    return wenyan.strip_inline_tags(text).strip()


def is_yilun_actual_exercise_start(text: str) -> bool:
    """识别真正带答案的试题起点，避免把方法步骤 1/2/3 误当题目。"""

    value = plain(text)
    if not re.match(r"^\s*\d+[\.．、]\s*", value):
        return False
    tail = re.sub(r"^\s*\d+[\.．、]\s*", "", value).strip()
    if re.match(r"^[（(]?\s*(?:20\d{2}|19\d{2})\s*[·.)）]", tail):
        return True
    return bool(
        re.match(
            r"^(?:阅读下面|下列对|下列关于|根据材料|结合材料|请|文中|把材料|简要|概括|分析|说明|指出|判断|比较)",
            tail,
        )
    )


def first_yilun_actual_exercise_index(lines: list[str]) -> int | None:
    for index, line in enumerate(lines):
        if is_yilun_actual_exercise_start(line):
            return index
    return None


def meaningful_line(text: str) -> bool:
    """判断一行是否应该进入 PPT；图片占位行即使无可见文字也要保留。"""

    return bool(plain(text)) or "<img" in text or "data-fbd-tif" in text


def compact(text: str) -> str:
    """用于识别标题；保留正文本身，不改变其可见排版内容。"""

    return re.sub(r"[\s_＿]+", "", plain(text))


def normalize_heading(text: str) -> str:
    value = plain(text)
    value = re.sub(r"[_＿]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def is_chinese_number_line(text: str) -> bool:
    return compact(text) in CHINESE_NUMBERS


def nonempty_indexes(lines: list[str], start: int = 0) -> list[int]:
    return [index for index in range(start, len(lines)) if plain(lines[index])]


def next_nonempty_index(lines: list[str], index: int) -> int | None:
    for candidate in range(index + 1, len(lines)):
        if plain(lines[candidate]):
            return candidate
    return None


def find_first_compact(lines: list[str], target: str, start: int = 0) -> int | None:
    target_compact = re.sub(r"[\s_＿]+", "", target)
    for index in range(start, len(lines)):
        if compact(lines[index]) == target_compact:
            return index
    return None


def collect_tif_keep_map(source_dir: Path) -> dict[str, str]:
    result: dict[str, str] = {}
    tif_paths = [
        path
        for path in sorted(source_dir.rglob("*"))
        if path.is_file() and path.suffix.lower() in {".tif", ".tiff"}
    ]
    for path in tif_paths:
        if path.is_file():
            result[path.name.lower()] = tif_markdown_tag(path.name)
    return result


def convert_yilun_fbd_to_md(
    *,
    body_fbd: Path = DEFAULT_BODY_FBD,
    answer_fbd: Path = DEFAULT_ANSWER_FBD,
    output_root: Path = DEFAULT_MD_ROOT,
    progress: Callable[[str], None] | None = None,
) -> Path:
    """将正文、答案 FBD 转成一个可重复使用的 MD run。"""

    def emit(message: str) -> None:
        (progress or print)(message)

    run_dir = output_root / datetime.now().strftime("run_%Y%m%d_%H%M%S")
    run_dir.mkdir(parents=True, exist_ok=False)
    tif_map = collect_tif_keep_map(body_fbd.parent)
    outputs: list[dict[str, object]] = []
    for role, path, name in (
        ("正文", body_fbd, "正文_body.md"),
        ("答案", answer_fbd, "答案_answer.md"),
    ):
        emit(f"转换{role} FBD：{path}")
        raw = read_fbd(path, "gb18030")
        markdown, stats = convert_fbd_to_markdown(
            raw,
            tif_replacements=tif_map,
            auto_start=False,
        )
        md_path = run_dir / name
        md_path.write_text(markdown, encoding="utf-8", newline="\n")
        outputs.append(
            {
                "role": role,
                "source": str(path),
                "markdown": str(md_path),
                "stats": asdict(stats),
            }
        )
        emit(f"{role} MD：{md_path}，TIF 保留 {stats.tif_kept} 个")
    (run_dir / "conversion_manifest.json").write_text(
        json.dumps(
            {
                "created_at": datetime.now().isoformat(timespec="seconds"),
                "outputs": outputs,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
        newline="\n",
    )
    return run_dir


def find_md_files(run_dir: Path) -> tuple[Path, Path]:
    body = sorted(run_dir.glob("正文_*.md"))
    answer = sorted(run_dir.glob("答案_*.md"))
    if not body:
        body = sorted(run_dir.glob("*body.md"))
    if not answer:
        answer = sorted(run_dir.glob("*answer.md"))
    if not body or not answer:
        raise FileNotFoundError(f"未找到正文和答案 MD：{run_dir}")
    return body[0], answer[0]


def slice_reading_one(md_path: Path) -> list[str]:
    """取出正文或答案中“阅读Ⅰ”到“阅读Ⅱ”之间的内容。"""

    lines = [line for line in wenyan.clean_lines(md_path.read_text(encoding="utf-8")) if line.strip()]
    start = find_first_compact(lines, TARGET_VOLUME)
    end = find_first_compact(lines, READING_II, (start or 0) + 1)
    if start is None:
        raise ValueError(f"{md_path.name} 中没有找到 {TARGET_VOLUME}")
    if end is None:
        end = len(lines)
    block = lines[start + 1 : end]
    # FBD 中首页标题和目录标题各出现一次，去掉重复的根标题，避免生成空白标题段。
    target_compact = compact(TARGET_VOLUME)
    return [line for line in block if compact(line) != target_compact]


def parse_topic_header(line: str) -> tuple[int, str] | None:
    value = normalize_heading(line)
    match = re.match(r"^专题([一二三四五六七八九十]+)\s*(.*)$", value)
    if not match:
        return None
    number = CHINESE_NUMBERS.get(match.group(1))
    if not number:
        return None
    return number, match.group(2).strip()


def parse_break_header(line: str) -> tuple[int, str] | None:
    value = normalize_heading(line)
    match = re.match(r"^突破(\d+)\s*(.*)$", value)
    if not match:
        return None
    return int(match.group(1)), match.group(2).strip()


def body_break_header_at(lines: list[str], index: int) -> tuple[int, str] | None:
    value = plain(lines[index])
    if not re.fullmatch(r"\d+", value):
        return None
    next_index = next_nonempty_index(lines, index)
    if next_index is None:
        return None
    next_value = plain(lines[next_index])
    if re.match(r"^\d+[\.．、]", next_value):
        return None
    if is_chinese_number_line(next_value):
        return None
    return int(value), normalize_heading(next_value)


def append_segment(
    segments: list[tuple[str, list[str], tuple[int, int] | None]],
    heading: str,
    lines: list[str],
    key: tuple[int, int] | None,
) -> None:
    cleaned = [line for line in lines if meaningful_line(line)]
    if cleaned:
        segments.append((heading, cleaned, key))


def table_key(text: str) -> str:
    return compact(text)


def yilun_tif_names(text: str) -> list[str]:
    return re.findall(r'data-fbd-tif="([^"]+)"', text, flags=re.I)


def is_yilun_large_tif_name(name: str) -> bool:
    value = Path(name).name.lower()
    return bool(re.fullmatch(r"a\d+\.tiff?", value) or value in {"b1a.tif"})


def has_yilun_large_tif(text: str) -> bool:
    return any(is_yilun_large_tif_name(name) for name in yilun_tif_names(text))


YILUN_TIF_TAG_RE = re.compile(r'<img\b[^>]*data-fbd-tif="([^"]+)"[^>]*?/?>', re.I)


def strip_yilun_large_tif_tags(text: str) -> tuple[str, str, list[str]]:
    """拆出正文里的大图 TIF。

    普通 FBD 图片标签是行内小图；`a1.tif` 这类是正文图表，需要独立
    占版面。这里只剥离大图标签，保留同一行剩余文字作为普通正文。
    """

    large_tags: list[str] = []
    large_names: list[str] = []

    def repl(match: re.Match[str]) -> str:
        name = match.group(1)
        if not is_yilun_large_tif_name(name):
            return match.group(0)
        large_tags.append(match.group(0))
        large_names.append(name)
        return ""

    cleaned = YILUN_TIF_TAG_RE.sub(repl, text)
    return cleaned.strip(), "".join(large_tags), large_names


def mark_yilun_large_tif_lines(lines: list[wenyan.StyledLine]) -> list[wenyan.StyledLine]:
    """把大图 TIF 行改为独立内容块并预留高度，避免后文盖住图片。"""

    result: list[wenyan.StyledLine] = []
    for line in lines:
        if (line.role or "") == "yilun_table" or not has_yilun_large_tif(line.text):
            result.append(line)
            continue

        cleaned, large_markup, large_names = strip_yilun_large_tif_tags(line.text)
        image_first = bool(re.match(r"^\s*<img\b", line.text, flags=re.I))
        image_line = wenyan.StyledLine(
            large_markup,
            "yilun_large_tif_line",
            {"large_tifs": large_names, "force_new_chunk": True},
        )
        spacer_lines = [
            wenyan.StyledLine("", "yilun_large_tif_spacer")
            for _ in range(YILUN_LARGE_TIF_SPACER_LINES)
        ]

        if image_first:
            result.append(image_line)
            result.extend(spacer_lines)
            if meaningful_line(cleaned):
                result.append(wenyan.StyledLine(cleaned, line.role, line.meta))
        else:
            if meaningful_line(cleaned):
                result.append(wenyan.StyledLine(cleaned, line.role, line.meta))
            result.append(image_line)
            result.extend(spacer_lines)
    return result


def prepare_yilun_styled_lines(lines: list[wenyan.StyledLine]) -> list[wenyan.StyledLine]:
    return mark_yilun_large_tif_lines(apply_yilun_tables_to_styled(lines))


def is_yilun_table_header(text: str) -> bool:
    key = table_key(text)
    return key in YILUN_TABLE_HEADER_KEYS or key in YILUN_DYNAMIC_TABLE_SPECS


def convert_yilun_fbd_table_cell(raw_cell: str, tif_map: dict[str, str]) -> str:
    """将 FBD 表格单元格转换成可直接放入 PPT 单元格的内联 Markdown。"""

    converted, _stats = convert_fbd_to_markdown(
        raw_cell,
        auto_start=False,
        include_css=False,
        tif_replacements=tif_map,
    )
    return converted.strip()


def extract_yilun_dynamic_table_specs(body_fbd: Path) -> dict[str, list[dict[str, object]]]:
    """从源 FBD 的 BG/BH 标记中恢复所有可识别的表格。

    正文 MD 是为了通用阅读而输出的线性文本，表格列分隔符会在公共
    转换器中被移除。这里直接读取原 FBD，只把表格的行列结构恢复出来，
    不修改单元格正文，也不依赖某一页或某一种题型。
    """

    if not body_fbd.exists():
        return {}
    tif_map = collect_tif_keep_map(body_fbd.parent)
    raw = read_fbd(body_fbd, "gb18030")
    result: dict[str, list[dict[str, object]]] = {}
    table_pattern = re.compile(
        r"〖BG(?!\))[^〗]*〗(?P<body>.*?)(?=〖BG\)F〗)",
        re.S,
    )
    row_pattern = re.compile(
        r"〖BH[^〗]*〗(?P<row>.*?)(?=〖BH[^〗]*〗|$)",
        re.S,
    )

    for source_index, table_match in enumerate(table_pattern.finditer(raw), 1):
        row_matches = list(row_pattern.finditer(table_match.group("body")))
        if len(row_matches) < 2:
            continue

        parsed_rows: list[list[str]] = []
        for row_match in row_matches:
            cells = [
                convert_yilun_fbd_table_cell(part, tif_map)
                for part in row_match.group("row").split("〖〗")
            ]
            while cells and not plain(cells[-1]):
                cells.pop()
            if any(plain(cell) for cell in cells):
                parsed_rows.append(cells)
        if len(parsed_rows) < 2:
            continue

        headers = [cell.strip() for cell in parsed_rows[0]]
        header_key = "".join(compact(cell) for cell in headers)
        if not header_key:
            continue
        column_count = len(headers)
        rows: list[list[str]] = []
        for row in parsed_rows[1:]:
            normalized = row[:column_count] + [""] * max(0, column_count - len(row))
            if any(plain(cell) for cell in normalized):
                rows.append(normalized)
        if not rows:
            continue

        result.setdefault(header_key, []).append(
            {
                "headers": headers,
                "rows": rows,
                "source_index": source_index,
            }
        )
    return result


def dynamic_yilun_table_spec_for(
    lines: list[wenyan.StyledLine],
    index: int,
    key: str,
) -> dict[str, object] | None:
    """从同名表头候选中选出与当前 MD 片段最匹配的原始表格。"""

    candidates = YILUN_DYNAMIC_TABLE_SPECS.get(key, [])
    if not candidates:
        return None
    following = "".join(
        compact(lines[cursor].text)
        for cursor in range(index + 1, min(len(lines), index + 90))
    )

    def score(spec: dict[str, object]) -> tuple[int, int]:
        rows = spec.get("rows", [])
        score_value = 0
        for row in rows[:4]:  # type: ignore[union-attr]
            for cell in row:  # type: ignore[union-attr]
                token = compact(str(cell))
                if len(token) >= 4 and token in following:
                    score_value += min(24, len(token))
        return score_value, -int(spec.get("source_index", 0))

    return max(candidates, key=score)


def consume_dynamic_yilun_table_lines(
    lines: list[wenyan.StyledLine],
    start: int,
    spec: dict[str, object],
) -> int:
    """找到动态表格在扁平化 MD 中的结束位置。"""

    rows = spec.get("rows", [])
    target = ""
    for row in reversed(rows):  # type: ignore[union-attr]
        for cell in reversed(row):  # type: ignore[union-attr]
            value = compact(str(cell))
            if len(value) >= 8:
                target = value
                break
        if target:
            break

    joined = ""
    cursor = start
    while cursor < len(lines):
        value = compact(lines[cursor].text)
        if cursor > start and is_yilun_table_header(value):
            break
        if cursor > start and "<img" in lines[cursor].text and joined:
            break
        joined += value
        if target and target in joined:
            return cursor + 1
        cursor += 1
    return cursor


def is_probable_yilun_heading(text: str) -> bool:
    value = plain(text)
    return bool(
        re.match(r"^[一二三四五六七八九十]+、", value)
        or re.match(r"^\([一二三四五六七八九十]+\)", value)
        or re.fullmatch(r"[一二三四五六七八九十]", value)
        or re.fullmatch(r"\d+", value)
    )


def split_label_from_row(text: str, row_re: str | None) -> tuple[str, str]:
    if not row_re:
        return "", text
    match = re.match(row_re, text)
    if not match:
        return "", text
    label = match.group(0)
    return label, text[len(label) :].strip()


def split_yilun_row_cells(headers: list[str], text: str, row_re: str | None) -> list[str]:
    label, rest = split_label_from_row(text, row_re)
    if len(headers) == 6:
        topic_match = re.search(
            r"^(?P<feature>.*?)(?P<topic>[^（(]+[（(]20\d{2}[^)）]+[)）])(?P<materials>.*)$",
            rest,
        )
        if topic_match:
            feature = topic_match.group("feature").strip()
            topic = topic_match.group("topic").strip()
            materials = topic_match.group("materials").strip()
            segment_pattern = re.compile(r"[（(][^（）()]*?(?:《[^》]+》|取材于[^（）()]+)[^（）()]*[)）]")
            matches = list(segment_pattern.finditer(materials))
            segments: list[str] = []
            cursor = 0
            for match in matches[:3]:
                segments.append(materials[cursor:match.end()].strip())
                cursor = match.end()
            tail = materials[cursor:].strip()
            if tail or len(segments) < 3:
                segments.append(tail)
            while len(segments) < 3:
                segments.append("")
            return [label or text, feature, topic, segments[0], segments[1], segments[2]]
        return [label or text, rest, "", "", "", ""]

    if len(headers) == 2:
        return [label or text, rest] if label else [text, ""]

    if len(headers) == 3:
        if not label:
            return [text, "", ""]
        if "警惕" in rest:
            before, after = rest.split("警惕", 1)
            return [label, before.strip(), f"警惕{after}".strip()]
        for marker in ("注意", "核对", "定位", "仔细", "紧盯"):
            if marker in rest and rest.index(marker) > 10:
                before, after = rest.split(marker, 1)
                return [label, before.strip(), f"{marker}{after}".strip()]
        for term in sorted(YILUN_ABILITY_TERMS, key=len, reverse=True):
            if rest.endswith(term) and len(rest) > len(term):
                return [label, rest[: -len(term)].strip(), term]
        split_at = max(10, int(len(rest) * 0.62))
        return [label, rest[:split_at].strip(), rest[split_at:].strip()]

    cells = [label or text]
    if label:
        cells.append(rest)
    while len(cells) < len(headers):
        cells.append("")
    return cells[: len(headers)]


def group_yilun_table_rows(spec: dict[str, object], raw_lines: list[wenyan.StyledLine]) -> list[list[str]]:
    row_re = spec.get("row_re")
    row_re = str(row_re) if row_re else None
    headers = [str(item) for item in spec["headers"]]  # type: ignore[index]
    mode = str(spec.get("mode", "rows"))
    texts = [line.text for line in raw_lines if plain(line.text)]
    if not texts:
        return []

    if mode == "first_rest":
        return [[texts[0], " ".join(texts[1:]).strip()]]

    groups: list[str] = []
    current = ""
    for text in texts:
        plain_text = plain(text)
        starts_row = bool(row_re and re.match(row_re, plain_text))
        if starts_row and current:
            groups.append(current.strip())
            current = plain_text
        else:
            current = f"{current} {plain_text}".strip() if current else plain_text
    if current:
        groups.append(current.strip())

    return [split_yilun_row_cells(headers, row, row_re) for row in groups]


def yilun_table_column_weights(headers: list[str], rows: list[list[str]] | None = None) -> list[float]:
    if len(headers) == 2:
        base = [0.24, 0.76]
    elif len(headers) == 3:
        base = [0.16, 0.54, 0.30]
    elif len(headers) == 4:
        base = [0.15, 0.31, 0.27, 0.27]
    elif len(headers) == 6:
        base = [0.09, 0.24, 0.16, 0.17, 0.17, 0.17]
    else:
        base = [1.0 / len(headers)] * len(headers)

    if not rows:
        return base

    demand: list[float] = []
    for col_index, header in enumerate(headers):
        samples = [plain(header)]
        for row in rows:
            if col_index < len(row):
                cell = plain(row[col_index])
                if cell:
                    samples.append(cell)
        units = [wenyan.text_units(sample) for sample in samples if sample]
        if not units:
            demand.append(1.0)
            continue
        peak = max(units)
        avg = sum(units) / len(units)
        demand.append(max(1.0, peak * 0.72 + avg * 0.28))

    total = sum(demand)
    content_weights = [value / total for value in demand]
    blend = 0.55 if len(headers) <= 3 else 0.45
    weights = [base[i] * (1 - blend) + content_weights[i] * blend for i in range(len(base))]

    minimums = {
        2: [0.22, 0.24],
        3: [0.15, 0.24, 0.18],
        4: [0.12, 0.16, 0.16, 0.16],
        6: [0.07, 0.16, 0.12, 0.12, 0.12, 0.12],
    }.get(len(headers))
    if minimums:
        weights = [max(weight, minimums[i]) for i, weight in enumerate(weights)]
    total = sum(weights)
    return [weight / total for weight in weights]


def yilun_table_row_height(cells: list[str], widths: list[float], total_width: float) -> float:
    max_lines = 1
    font_size = 14.0
    for text, weight in zip(cells, widths):
        col_width = max(0.8, total_width * weight)
        # Chinese glyphs in PowerPoint occupy close to one full em.  The
        # conservative factor leaves room for punctuation and cell padding.
        capacity = max(6.0, col_width * 72.0 / font_size * 0.70)
        max_lines = max(max_lines, math.ceil(wenyan.text_units(plain(text)) / capacity))
    line_height = font_size / 72.0 * 1.38
    return max(0.34, max_lines * line_height + 0.10)


def split_yilun_table_pages(headers: list[str], rows: list[list[str]], *, max_height: float = 5.35) -> list[list[list[str]]]:
    widths = yilun_table_column_weights(headers, rows)
    pages: list[list[list[str]]] = []
    current: list[list[str]] = []
    used = 0.42
    for row in rows:
        height = yilun_table_row_height(row, widths, 11.85)
        if current and used + height > max_height:
            pages.append(current)
            current = []
            used = 0.42
        current.append(row)
        used += height
    if current:
        pages.append(current)
    return pages


def table_lines_from_meta(
    headers: list[str],
    rows: list[list[str]],
    title: str,
    *,
    max_height: float = 5.35,
) -> list[wenyan.StyledLine]:
    result: list[wenyan.StyledLine] = []
    for page_index, page_rows in enumerate(
        split_yilun_table_pages(headers, rows, max_height=max_height),
        1,
    ):
        result.append(
            wenyan.StyledLine(
                "",
                "yilun_table",
                {
                    "headers": headers,
                    "rows": page_rows,
                    "title": title,
                    "page_index": page_index,
                },
            )
        )
    return result


def apply_yilun_tables_to_styled(lines: list[wenyan.StyledLine]) -> list[wenyan.StyledLine]:
    """把 FBD 转 MD 后塌缩的表格块恢复为可编辑 PPT 表格。"""

    result: list[wenyan.StyledLine] = []
    index = 0
    while index < len(lines):
        current = lines[index]
        key = table_key(current.text)
        dynamic_spec = dynamic_yilun_table_spec_for(lines, index, key)
        spec = dynamic_spec if dynamic_spec is not None else YILUN_TABLE_SPECS.get(key)
        if spec is None:
            result.append(current)
            index += 1
            continue

        headers = [str(item) for item in spec["headers"]]  # type: ignore[index]
        if dynamic_spec is not None:
            rows = [
                [str(cell) for cell in row]
                for row in dynamic_spec.get("rows", [])  # type: ignore[union-attr]
            ]
            result.extend(
                table_lines_from_meta(
                    headers,
                    rows,
                    plain(current.text),
                )
            )
            index = consume_dynamic_yilun_table_lines(lines, index + 1, dynamic_spec)
            continue

        skip_next_header = str(spec.get("skip_next_header", ""))
        cursor = index + 1
        if skip_next_header and cursor < len(lines) and table_key(lines[cursor].text) == skip_next_header:
            cursor += 1

        raw_rows: list[wenyan.StyledLine] = []
        stop_re = str(spec.get("stop", ""))
        while cursor < len(lines):
            text = plain(lines[cursor].text)
            if not text:
                cursor += 1
                continue
            if is_yilun_table_header(text) and raw_rows:
                break
            if stop_re and re.match(stop_re, text) and raw_rows:
                break
            if raw_rows and is_probable_yilun_heading(text) and not spec.get("row_re"):
                break
            raw_rows.append(lines[cursor])
            cursor += 1

        rows = group_yilun_table_rows(spec, raw_rows)
        if not rows:
            result.append(current)
            index += 1
            continue

        result.extend(table_lines_from_meta(headers, rows, plain(current.text)))
        index = cursor
    return result


def split_yilun_body_segments(block: list[str]) -> list[tuple[str, list[str], tuple[int, int] | None]]:
    """把正文拆成可独立分页、可独立匹配答案的专题/突破片段。"""

    topic_breakthrough = find_first_compact(block, "专题突破")
    if topic_breakthrough is None:
        return [("阅读Ⅰ 信息类文本阅读", block, None)]

    segments: list[tuple[str, list[str], tuple[int, int] | None]] = []
    board_start = find_first_compact(block, "板块必备")
    board_start = board_start if board_start is not None else 0
    append_segment(segments, "板块必备", block[board_start + 1 : topic_breakthrough], None)

    index = topic_breakthrough + 1
    while index < len(block):
        if not is_chinese_number_line(block[index]):
            index += 1
            continue
        topic_number = CHINESE_NUMBERS[compact(block[index])]
        if topic_number > 2:
            break
        title_index = next_nonempty_index(block, index)
        if title_index is None:
            break
        topic_title = normalize_heading(block[title_index])
        next_topic = title_index + 1
        while next_topic < len(block) and not is_chinese_number_line(block[next_topic]):
            next_topic += 1
        topic_content = block[title_index + 1 : next_topic]
        topic_heading = f"专题{compact(block[index])} {topic_title}"

        break_positions: list[tuple[int, int, str]] = []
        cursor = 0
        while cursor < len(topic_content):
            header = body_break_header_at(topic_content, cursor)
            if header:
                break_number, break_title = header
                title_pos = next_nonempty_index(topic_content, cursor)
                if title_pos is not None:
                    break_positions.append((cursor, break_number, break_title))
                    cursor = title_pos + 1
                    continue
            cursor += 1

        if not break_positions:
            append_segment(segments, topic_heading, topic_content, None)
        else:
            first_start = break_positions[0][0]
            append_segment(segments, topic_heading, topic_content[:first_start], None)
            for position, (start, break_number, break_title) in enumerate(break_positions):
                title_pos = next_nonempty_index(topic_content, start)
                if title_pos is None:
                    continue
                end = break_positions[position + 1][0] if position + 1 < len(break_positions) else len(topic_content)
                content = topic_content[title_pos + 1 : end]
                heading = f"突破{break_number} {break_title}"
                append_segment(segments, heading, content, (topic_number, break_number))
        index = next_topic
    return segments


def split_yilun_answer_sections(block: list[str]) -> dict[tuple[int, int], list[str]]:
    """答案以专题/突破为边界切片，解决每个突破都从第 1 题重新编号的问题。"""

    result: dict[tuple[int, int], list[str]] = {}
    current_topic: int | None = None
    current_break: int | None = None
    for line in block:
        topic = parse_topic_header(line)
        if topic:
            current_topic = topic[0]
            current_break = None
            continue
        breakthrough = parse_break_header(line)
        if breakthrough and current_topic is not None:
            current_break = breakthrough[0]
            result[(current_topic, current_break)] = []
            continue
        if current_topic is not None and current_break is not None:
            result[(current_topic, current_break)].append(line)
    return result


def drop_leading_yilun_titles(lines: list[str], titles: tuple[str, ...]) -> list[str]:
    title_keys = {compact(title) for title in titles}
    index = 0
    while index < len(lines) and compact(lines[index]) in title_keys:
        index += 1
    return lines[index:]


def slice_yilun_block(
    lines: list[str],
    start_marker: str,
    end_marker: str | None,
    strip_titles: tuple[str, ...],
) -> list[str]:
    start = find_first_compact(lines, start_marker)
    if start is None:
        return []
    end = len(lines)
    if end_marker:
        candidate = find_first_compact(lines, end_marker, start + 1)
        if candidate is not None:
            end = candidate
    block = lines[start:end]
    return drop_leading_yilun_titles(block, strip_titles)


def split_yilun_board_segments(block: list[str]) -> list[tuple[str, list[str]]]:
    """把一个板块切成“板块必备 + 各专题”。

    这里不限制专题数量，适用于整本书的各个大板块。
    """

    segments: list[tuple[str, list[str]]] = []
    board_start = find_first_compact(block, "板块必备")
    topic_breakthrough = find_first_compact(block, "专题突破")
    if board_start is not None and topic_breakthrough is not None and board_start < topic_breakthrough:
        append_segment(segments, "板块必备", block[board_start + 1 : topic_breakthrough], None)

    def topic_header_at(position: int) -> tuple[str, int] | None:
        """兼容正文的“专题一\n标题”和答案的“专题一____标题”。"""

        line = block[position]
        if is_chinese_number_line(line):
            title_index = next_nonempty_index(block, position)
            if title_index is None:
                return None
            return normalize_heading(block[title_index]), title_index + 1
        topic = parse_topic_header(line)
        if topic:
            title = topic[1] or normalize_heading(line)
            return title, position + 1
        return None

    index = (topic_breakthrough + 1) if topic_breakthrough is not None else 0
    while index < len(block):
        header = topic_header_at(index)
        if not header:
            index += 1
            continue
        topic_title, content_start = header
        next_topic = content_start
        while next_topic < len(block) and topic_header_at(next_topic) is None:
            next_topic += 1
        topic_content = block[content_start:next_topic]
        append_segment(segments, topic_title, topic_content, None)
        index = next_topic
    return segments


def parse_yilun_board_pieces_from_block(
    block: list[str],
    *,
    volume: str,
    num_offset: int = 0,
) -> list[wenyan.Piece]:
    pieces: list[wenyan.Piece] = []
    for index, (heading, content, _key) in enumerate(split_yilun_board_segments(block), 1):
        cleaned = [line for line in content if meaningful_line(line)]
        if not cleaned:
            continue
        pieces.append(
            wenyan.Piece(
                volume=volume,
                num=index + num_offset,
                title=heading,
                author="",
                lines=[heading, MAIN_MARKER, *cleaned],
            )
        )
    return pieces


def pair_yilun_body_and_answer_pieces(
    body_pieces: list[wenyan.Piece],
    answer_pieces: list[wenyan.Piece],
) -> dict[tuple[str, int], wenyan.Piece]:
    if not body_pieces or not answer_pieces:
        return {}
    offset = 0
    if body_pieces[0].title == "板块必备" and answer_pieces[0].title != "板块必备":
        offset = 1
    answer_map: dict[tuple[str, int], wenyan.Piece] = {}
    for body_index, body_piece in enumerate(body_pieces):
        answer_index = body_index - offset
        if answer_index < 0 or answer_index >= len(answer_pieces):
            continue
        answer_map[wenyan.piece_key(body_piece)] = answer_pieces[answer_index]
    return answer_map


def build_yilun_board_cache(body_md: Path, answer_md: Path) -> dict[str, tuple[list[wenyan.Piece], dict[tuple[str, int], wenyan.Piece]]]:
    body_lines = [line for line in wenyan.clean_lines(body_md.read_text(encoding="utf-8")) if line.strip()]
    answer_lines = [line for line in wenyan.clean_lines(answer_md.read_text(encoding="utf-8")) if line.strip()]
    result: dict[str, tuple[list[wenyan.Piece], dict[tuple[str, int], wenyan.Piece]]] = {}
    for spec in YILUN_BOARD_SPECS:
        body_block = slice_yilun_block(
            body_lines,
            str(spec["start"]),
            str(spec["end"]) if spec["end"] is not None else None,
            tuple(spec["strip_titles"]),  # type: ignore[arg-type]
        )
        answer_block = slice_yilun_block(
            answer_lines,
            str(spec["start"]),
            str(spec["end"]) if spec["end"] is not None else None,
            tuple(spec["strip_titles"]),  # type: ignore[arg-type]
        )
        body_pieces = parse_yilun_board_pieces_from_block(
            body_block,
            volume=str(spec["volume"]),
            num_offset=int(spec.get("num_offset", 0)),
        )
        answer_pieces = parse_yilun_board_pieces_from_block(
            answer_block,
            volume=str(spec["volume"]),
            num_offset=int(spec.get("num_offset", 0)),
        )
        result[str(spec["key"])] = (body_pieces, pair_yilun_body_and_answer_pieces(body_pieces, answer_pieces))
    return result


def build_yilun_pieces(body_md: Path, answer_md: Path) -> tuple[list[wenyan.Piece], dict[tuple[str, int], wenyan.Piece]]:
    body_block = slice_reading_one(body_md)
    answer_block = slice_reading_one(answer_md)
    answer_sections = split_yilun_answer_sections(answer_block)
    pieces: list[wenyan.Piece] = []
    answer_map: dict[tuple[str, int], wenyan.Piece] = {}
    topic_volumes: dict[int, str] = {}

    for num, (heading, content, answer_key) in enumerate(split_yilun_body_segments(body_block), 1):
        topic = parse_topic_header(heading)
        if topic:
            topic_volumes[topic[0]] = heading
            volume = heading
        elif answer_key:
            volume = topic_volumes.get(answer_key[0], f"专题{answer_key[0]}")
        else:
            volume = heading

        piece = wenyan.Piece(
            volume=volume,
            num=num,
            title=heading,
            author="",
            lines=[heading, MAIN_MARKER, *content],
        )
        pieces.append(piece)

        if answer_key:
            segment_answer = answer_sections.get(answer_key, [])
            if segment_answer:
                answer_piece = wenyan.Piece(
                    volume=volume,
                    num=num,
                    title=heading,
                    author="",
                    lines=[heading, MAIN_MARKER, *segment_answer],
                )
                answer_map[wenyan.piece_key(piece)] = answer_piece
    return pieces, answer_map


def style_segment(heading: str, content: list[str], answer_lines: list[str] | None) -> list[wenyan.StyledLine]:
    heading_line = wenyan.StyledLine(
        heading,
        "board_main_heading" if heading == "板块必备" else "heading",
    )
    if heading in {"板块必备", "专题突破"}:
        if heading == "板块必备":
            return [heading_line, *style_yilun_board_lines(content)]
        return [heading_line, *prepare_yilun_styled_lines(wenyan.styled_body_lines(content, heading))]
    if is_yilun_judgement_group(content):
        styled = build_yilun_judgement_section(content, answer_lines)
        return [heading_line, *prepare_yilun_styled_lines(styled)]
    if any("正误关系" in plain(line) for line in content):
        content = remove_yilun_judgement_tail_blank(content)

    exercise_start = first_yilun_actual_exercise_index(content) if answer_lines else None
    if answer_lines and exercise_start is None:
        styled = yuedu.style_yuedu_styled_lines(wenyan.styled_body_lines(content, heading))
        return [heading_line, *prepare_yilun_styled_lines(styled)]
    if answer_lines and exercise_start and exercise_start > 0:
        intro = content[:exercise_start]
        exercise = content[exercise_start:]
        intro_styled = yuedu.style_yuedu_styled_lines(wenyan.styled_body_lines(intro, heading))
        exercise_styled, _matched = yuedu.build_yuedu_styled_section(exercise, answer_lines)
        exercise_styled = yuedu.style_yuedu_styled_lines(exercise_styled)
        return [heading_line, *prepare_yilun_styled_lines([*intro_styled, *exercise_styled])]

    styled, _matched = yuedu.build_yuedu_styled_section(content, answer_lines)
    styled = yuedu.style_yuedu_styled_lines(styled)
    return [heading_line, *prepare_yilun_styled_lines(styled)]


def normalize_yilun_two_part_judgement_answers(
    items: dict[int, wenyan.AnswerItem],
) -> dict[int, wenyan.AnswerItem]:
    """拆分“正误关系/设错角度及理由”答案中的两个小空。

    这本书的答案 MD 会把这类答案写成：
    ``(1)错误<span class="blank">...</span>(2)逻辑错乱。理由……``
    公共解析器原本只按下划线字符切分，因此 HTML blank 标签会把两项
    粘成一个答案。这里只处理首项明确为“错误/正确”的这种判断题，
    避免影响翻译题等其他带 (1)/(2) 的题型。
    """

    parenthetical_re = re.compile(r"[（(]\s*(\d+)\s*[）)]")
    for item in items.values():
        if not item.answers:
            continue
        first_value = ""
        second_raw = ""

        # clean_lines() turns answer blank spans into underscores, so the
        # common parser normally returns two pieces: ``(1)错误`` and
        # ``(2)逻辑错乱。理由……``.
        if len(item.answers) >= 2:
            first_match = re.match(r"^[（(]\s*1\s*[）)]\s*(.*)$", item.answers[0])
            second_match = re.match(r"^[（(]\s*2\s*[）)]\s*(.*)$", item.answers[1])
            if first_match and second_match:
                first_value = plain(first_match.group(1))
                second_raw = second_match.group(1)

        # Keep support for an unnormalized one-piece answer when this
        # adapter is called with raw lines rather than clean_lines output.
        if not first_value and len(item.answers) == 1:
            raw = item.answers[0]
            labels = list(parenthetical_re.finditer(raw))
            if len(labels) >= 2 and labels[0].group(1) == "1" and labels[1].group(1) == "2":
                first_value = plain(raw[labels[0].end() : labels[1].start()])
                second_raw = raw[labels[1].end() :]

        first_value = re.sub(r"^[：:]\s*", "", first_value).strip()
        if first_value not in {"错误", "正确"} or not second_raw:
            continue

        second_plain = plain(second_raw)
        if not second_plain:
            continue

        # The second blank is an integrated answer area.  Its answer starts
        # with the judgement angle and continues through the complete reason;
        # do not shorten it to only the first phrase or move the tail into
        # an analysis box.
        second_value = second_plain.strip()
        explanation = ""
        if not second_value:
            continue

        item.answers = [first_value, second_value]
        if explanation:
            item.analyses.append(explanation)
    return items


def split_yilun_judgement_answer(answer: str) -> list[str]:
    """把判断题第二空拆成贴近横线宽度的答案行。"""

    text = re.sub(r"\s+", "", plain(answer))
    return [text] if text else []


def yilun_inline_underlined_answer(answer: str, target_units: float | None = None) -> str:
    text = f"　{plain(answer)}　"
    if target_units is not None:
        pad_units = max(2.0, target_units - wenyan.text_units(text))
        text = f"{text}{NBSP * max(4, int(math.ceil(pad_units / 0.55)))}"
    return wenyan.answer_span(text, underline=True)


def yilun_visible_underline(target_units: float) -> str:
    return f'<span class="uline">{NBSP * max(5, int(math.ceil(target_units)))}</span>'


def yilun_prefix_padding(units: float) -> str:
    return NBSP * max(0, int(math.ceil(units * 1.75)))


def yilun_judgement_answer_lines(
    answer: str,
    *,
    continuation_label: str | None = None,
) -> list[wenyan.StyledLine]:
    pieces = split_yilun_judgement_answer(answer)
    if not pieces:
        return []
    result: list[wenyan.StyledLine] = []
    for index, piece in enumerate(pieces):
        meta = {
            "judgement_answer": piece,
            "judgement_prefix_units": 0.0,
            "judgement_answer_group": continuation_label or "",
        }
        if index > 0 and continuation_label:
            meta["judgement_continuation_label"] = continuation_label
        result.append(
            wenyan.StyledLine(
                wenyan.answer_span(piece, underline=True, pad_underline=True),
                "yilun_judgement_answer",
                meta,
            )
        )
    return result


def repair_yilun_judgement_continuation_chunks(
    chunks: list[list[wenyan.StyledLine]],
) -> list[list[wenyan.StyledLine]]:
    repaired: list[list[wenyan.StyledLine]] = []
    for chunk in chunks:
        current: list[wenyan.StyledLine] = []
        seen_judgement_answer = False
        for line in chunk:
            line_text = plain(line.text)
            starts_next_numbered_question = wenyan.is_numbered_question_line(line)
            starts_next_subitem = bool(re.match(r"^\s*[（(]\s*\d+\s*[）)]", line_text)) and not (
                "设错角度及理由" in line_text
                or "验证技巧" in line_text
                or "设误点及分析" in line_text
            )
            if current and seen_judgement_answer and (starts_next_numbered_question or starts_next_subitem):
                repaired.append(current)
                current = []
                seen_judgement_answer = False
            current.append(line)
            if line.meta and line.meta.get("judgement_answer"):
                seen_judgement_answer = True
        if current:
            repaired.append(current)

    merged: list[list[wenyan.StyledLine]] = []
    for chunk in repaired:
        starts_with_continuation = bool(
            chunk
            and chunk[0].meta
            and chunk[0].meta.get("judgement_continuation_label")
        )
        answer_only_chunk = bool(
            chunk
            and all((line.role or "") in {"yilun_judgement_answer", "answer_spacer"} for line in chunk)
        )
        if (
            (starts_with_continuation or answer_only_chunk)
            and merged
            and sum(wenyan.line_page_cost(line) for line in [*merged[-1], *chunk])
            <= YILUN_JUDGEMENT_MAX_PAGE_UNITS
        ):
            merged[-1].extend(chunk)
            continue
        merged.append(chunk)

    fitted: list[list[wenyan.StyledLine]] = []
    for chunk in merged:
        first_answer = next(
            (index for index, line in enumerate(chunk) if (line.role or "") == "yilun_judgement_answer"),
            None,
        )
        if first_answer and first_answer > 0:
            lead_cost = sum(wenyan.line_page_cost(line) for line in chunk[:first_answer])
            answer_count = sum(1 for line in chunk[first_answer:] if (line.role or "") == "yilun_judgement_answer")
            if lead_cost >= 6 and answer_count >= 2:
                fitted.append(chunk[:first_answer])
                fitted.append(chunk[first_answer:])
                continue
        fitted.append(chunk)
    return fitted


def is_yilun_triple_judgement_question(lines: list[str]) -> bool:
    text = "\n".join(plain(line) for line in lines)
    return "出题方式" in text and "正误关系" in text and "设误点及分析" in text


def is_yilun_judgement_group(lines: list[str]) -> bool:
    text = "\n".join(plain(line) for line in lines)
    return "正误关系" in text and (
        "设错角度及理由" in text
        or "验证技巧" in text
        or "设误点及分析" in text
    )


def split_yilun_triple_judgement_answers(
    item: wenyan.AnswerItem | None,
) -> dict[int, tuple[str, str, str]]:
    if item is None:
        return {}
    answers, _analyses = yuedu.split_item_answer_analysis(item)
    result: dict[int, tuple[str, str, str]] = {}
    for index in range(0, len(answers), 3):
        group = answers[index : index + 3]
        if len(group) < 3:
            break
        result[len(result) + 1] = (plain(group[0]), plain(group[1]), plain(group[2]))
    return result


def build_yilun_triple_judgement_group(
    group_lines: list[str],
    item: wenyan.AnswerItem | None,
    section_title: str,
) -> list[wenyan.StyledLine]:
    triples = split_yilun_triple_judgement_answers(item)
    styled: list[wenyan.StyledLine] = []
    current_subnum: int | None = None
    in_original_material = False

    for line in group_lines:
        value = plain(line)
        if not value or yuedu.is_blank_only_answer_line(line):
            continue

        sub_match = re.match(r"^\s*[（(]\s*(\d+)\s*[）)]", value)
        if sub_match:
            current_subnum = int(sub_match.group(1))
            in_original_material = False

        if value.startswith("【原文】"):
            in_original_material = True
            styled.append(wenyan.StyledLine(line, "reading"))
            continue

        if re.match(r"^[①②③]", value):
            in_original_material = False
            answers = triples.get(current_subnum or 0)
            label = wenyan.strip_fillable_blanks_from_line(line)
            if not answers:
                styled.extend(wenyan.styled_body_lines([label], section_title))
                continue

            if value.startswith("①"):
                answer = answers[0]
                styled.extend(
                    wenyan.styled_body_lines(
                        [f"{label}{wenyan.answer_span(answer, underline=True, pad_underline=True)}"],
                        section_title,
                    )
                )
                continue
            if value.startswith("②"):
                answer = answers[1]
                styled.extend(
                    wenyan.styled_body_lines(
                        [f"{label}{wenyan.answer_span(answer, underline=True, pad_underline=True)}"],
                        section_title,
                    )
                )
                continue
            if value.startswith("③"):
                styled.extend(wenyan.styled_body_lines([label], section_title))
                styled.extend(yilun_judgement_answer_lines(answers[2], continuation_label=plain(label)))
                continue

        if in_original_material:
            styled.append(wenyan.StyledLine(line, "reading"))
        else:
            styled.extend(wenyan.styled_body_lines([line], section_title))
    return styled


def build_yilun_fallback_group(
    group_lines: list[str],
    answer_lines: list[str] | None,
    num: int,
) -> list[wenyan.StyledLine]:
    answer_block = yuedu.extract_numbered_answer_block(answer_lines, num)
    scoped_answers: list[str] | None = None
    if answer_block:
        scoped_answers = [f"{num}.{answer_block[0]}", *answer_block[1:]]
    fallback, _matched = yuedu.build_yuedu_styled_section(group_lines, scoped_answers)
    return yuedu.style_yuedu_styled_lines(fallback)


def mark_yilun_judgement_text_lines(lines: list[wenyan.StyledLine]) -> list[wenyan.StyledLine]:
    marked: list[wenyan.StyledLine] = []
    for line in lines:
        if (line.role or "") == "yilun_judgement_answer":
            marked.append(line)
            continue
        meta = dict(line.meta or {})
        meta["yilun_judgement_text"] = True
        marked.append(wenyan.StyledLine(line.text, line.role, meta))
    return marked


def build_yilun_judgement_section(
    lines: list[str],
    answer_lines: list[str] | None,
) -> list[wenyan.StyledLine]:
    """渲染本书“正误关系/设错角度及理由”两空题。

    这类题的第二空不是一个短词，而是“设错角度 + 完整理由”的整体答案。
    因此不能把两个答案拼成一个普通开放题答案，也不能把第二空限制在题干
    原本很短的 blank 标签里。第二小题保留题目标签，答案另起长答题行，
    让下划线和答案随着完整文本自然换行。
    """

    answer_items = wenyan.extract_answer_items(answer_lines)
    leading, groups = wenyan.split_body_numbered_groups(lines)
    section_title = lines[0] if lines else ""
    preamble_lines, _preamble_filled = wenyan.fill_lines_with_answer_values(
        leading,
        wenyan.extract_preamble_answers(answer_lines),
    )
    styled: list[wenyan.StyledLine] = wenyan.styled_body_lines(preamble_lines, section_title)

    for num, group_lines in groups:
        item = answer_items.get(num)
        if item is None:
            styled.extend(wenyan.styled_body_lines(group_lines, section_title))
            continue

        if is_yilun_triple_judgement_question(group_lines):
            styled.extend(build_yilun_triple_judgement_group(group_lines, item, section_title))
            continue

        group_text = "\n".join(plain(line) for line in group_lines)
        if "正误关系" not in group_text or not (
            "设错角度及理由" in group_text or "验证技巧" in group_text
        ):
            styled.extend(build_yilun_fallback_group(group_lines, answer_lines, num))
            continue

        answers, analyses = yuedu.split_item_answer_analysis(item)
        first_answer = answers[0] if answers else ""
        second_answer = answers[1] if len(answers) >= 2 else ""
        first_done = False
        second_done = False
        in_original_material = False

        for line in group_lines:
            value = plain(line)
            if not value or yuedu.is_blank_only_answer_line(line):
                continue

            # 判断题中的【原文】标签保留宋体，标签后的材料使用楷体；
            # 材料可能与标签处在同一行，因此用内联字体标签处理。
            marker = "【原文】"
            if marker in value:
                in_original_material = True
                marker_index = line.find(marker)
                if marker_index >= 0:
                    line = (
                        line[:marker_index]
                        + '<span class="font-song">【原文】</span>'
                        + line[marker_index + len(marker) :]
                    )
                styled.append(wenyan.StyledLine(line, "reading"))
                continue

            if in_original_material and re.match(r"^\s*[（(]\s*[12]\s*[）)]", value):
                in_original_material = False

            if in_original_material:
                styled.append(wenyan.StyledLine(line, "reading"))
                continue

            if re.match(r"^\s*[（(]\s*1\s*[）)]", value) and "正误关系" in value:
                label = wenyan.strip_fillable_blanks_from_line(line)
                if first_answer:
                    styled.extend(
                        wenyan.styled_body_lines(
                            [f"{label}{wenyan.answer_span(first_answer, underline=True, pad_underline=True)}"],
                            section_title,
                        )
                    )
                    first_done = True
                else:
                    styled.extend(wenyan.styled_body_lines([label], section_title))
                continue
            if re.match(r"^\s*[（(]\s*2\s*[）)]", value) and (
                "设错角度及理由" in value or "验证技巧" in value
            ):
                label = wenyan.strip_fillable_blanks_from_line(line)
                if second_answer:
                    styled.extend(wenyan.styled_body_lines([label], section_title))
                    styled.extend(yilun_judgement_answer_lines(second_answer, continuation_label=plain(label)))
                    second_done = True
                else:
                    styled.extend(wenyan.styled_body_lines([label], section_title))
                continue
            styled.extend(wenyan.styled_body_lines([line], section_title))

        # Keep any extra analysis attached to the same question, preserving the
        # existing red animated answer/analysis treatment.
        for analysis in analyses:
            styled.append(wenyan.StyledLine(analysis, "analysis"))

        # If an unusual answer file has only one parsed answer, still leave the
        # normal question text visible; do not fabricate a second answer.
        _ = first_done, second_done

    return mark_yilun_judgement_text_lines(styled)


def remove_yilun_judgement_tail_blank(lines: list[str]) -> list[str]:
    """移除判断题两小问之后遗留的独立空白行。

    该空白行在正文 FBD 中是版式占位，不是第三个答案位置；若保留，
    阅读类通用逻辑会把整题误判成开放性长答案。
    """

    result: list[str] = []
    for line in lines:
        if (
            result
            and plain(line).strip()
            and re.fullmatch(r"[_＿\s]+", plain(line).strip())
            and (
                "设错角度及理由" in plain(result[-1])
                or "验证技巧" in plain(result[-1])
            )
        ):
            continue
        result.append(line)
    return result


def combine_yilun_board_headings(lines: list[str]) -> list[wenyan.StyledLine]:
    result: list[wenyan.StyledLine] = []
    index = 0
    while index < len(lines):
        value = plain(lines[index])
        if value in CHINESE_NUMBERS and index + 1 < len(lines):
            next_value = plain(lines[index + 1])
            result.append(wenyan.StyledLine(f"{value} {next_value}", "board_heading"))
            index += 2
            continue
        result.append(wenyan.StyledLine(lines[index], "body"))
        index += 1
    return result


def board_section_heading(text: str) -> bool:
    return bool(re.match(r"^[一二三四五六七八九十]\s+", plain(text)))


def board_subsection_heading(text: str) -> bool:
    return bool(re.match(r"^[（(][一二三四五六七八九十]+[）)]", plain(text)))


def board_numbered_title(text: str) -> bool:
    return bool(re.match(r"^\d+[\.．、]\s*[^\d]", plain(text)))


def board_table_stop(text: str) -> bool:
    value = plain(text)
    return board_section_heading(value) or board_subsection_heading(value) or board_numbered_title(value)


def make_yilun_course_standard_table(lines: list[wenyan.StyledLine]) -> list[wenyan.StyledLine]:
    visible = [line for line in lines if plain(line.text)]
    if not visible:
        return []
    header = plain(visible[0].text)
    rows = [[plain(line.text)] for line in visible[1:]]
    return table_lines_from_meta([header], rows, header)


def extract_yilun_board_table(
    lines: list[wenyan.StyledLine],
    index: int,
    *,
    max_height: float,
) -> tuple[list[wenyan.StyledLine], int] | None:
    header_key = table_key(lines[index].text)
    spec = YILUN_TABLE_SPECS.get(header_key)
    if spec is None:
        return None
    cursor = index + 1
    raw_rows: list[wenyan.StyledLine] = []
    stop_re = str(spec.get("stop", ""))
    while cursor < len(lines):
        value = plain(lines[cursor].text)
        if not value:
            cursor += 1
            continue
        if board_table_stop(value) and raw_rows:
            break
        if stop_re and re.match(stop_re, value) and raw_rows:
            break
        if is_yilun_table_header(value) and raw_rows:
            break
        raw_rows.append(lines[cursor])
        cursor += 1
    headers = [str(item) for item in spec["headers"]]  # type: ignore[index]
    rows = group_yilun_table_rows(spec, raw_rows)
    return table_lines_from_meta(headers, rows, plain(lines[index].text), max_height=max_height), cursor


def style_yilun_board_lines(content: list[str]) -> list[wenyan.StyledLine]:
    """将板块必备中的标题和塌缩表格整理成可成组分页的结构。"""

    source = combine_yilun_board_headings([line for line in content if meaningful_line(line)])
    result: list[wenyan.StyledLine] = []
    index = 0
    while index < len(source):
        line = source[index]
        value = plain(line.text)

        if board_section_heading(value):
            result.append(line)
            index += 1
            continue
        if board_subsection_heading(value):
            result.append(wenyan.StyledLine(line.text, "board_subheading"))
            index += 1
            continue
        if table_key(value) == "2025年修订版新课标要求":
            result.append(line)
            index += 1
            continue
        if is_yilun_table_header(value):
            result.append(line)
            index += 1
            continue
        if board_numbered_title(value):
            result.append(wenyan.StyledLine(line.text, "board_numbered_title"))
            index += 1
            continue
        result.append(wenyan.StyledLine(line.text, "board_body"))
        index += 1

    # Second pass: replace each board table header with a table object.  The
    # course-standard table is a one-column special case whose first row is
    # the centered title and the following rows are the three requirements.
    output: list[wenyan.StyledLine] = []
    index = 0
    while index < len(result):
        current = result[index]
        key = table_key(current.text)
        if key == "2025年修订版新课标要求":
            cursor = index + 1
            rows: list[list[str]] = []
            while cursor < len(result) and not board_section_heading(result[cursor].text):
                value = plain(result[cursor].text)
                if value:
                    rows.append([value])
                cursor += 1
            output.extend(
                table_lines_from_meta(
                    [key],
                    rows,
                    key,
                    max_height=4.65,
                )
            )
            index = cursor
            continue
        dynamic_spec = dynamic_yilun_table_spec_for(result, index, key)
        if dynamic_spec is not None:
            headers = [str(item) for item in dynamic_spec["headers"]]  # type: ignore[index]
            rows = [
                [str(cell) for cell in row]
                for row in dynamic_spec.get("rows", [])  # type: ignore[union-attr]
            ]
            output.extend(table_lines_from_meta(headers, rows, key))
            index = consume_dynamic_yilun_table_lines(result, index + 1, dynamic_spec)
            continue
        if key in YILUN_TABLE_SPECS:
            spec = YILUN_TABLE_SPECS[key]
            cursor = index + 1
            raw_rows: list[wenyan.StyledLine] = []
            stop_re = str(spec.get("stop", ""))
            while cursor < len(result):
                value = plain(result[cursor].text)
                if not value:
                    cursor += 1
                    continue
                if board_table_stop(value) and raw_rows:
                    break
                if stop_re and re.match(stop_re, value) and raw_rows:
                    break
                if is_yilun_table_header(value) and raw_rows:
                    break
                raw_rows.append(result[cursor])
                cursor += 1
            headers = [str(item) for item in spec["headers"]]  # type: ignore[index]
            rows = group_yilun_table_rows(spec, raw_rows)
            output.extend(table_lines_from_meta(headers, rows, key, max_height=4.65))
            index = cursor
            continue
        if (current.role or "") not in {
            "board_main_heading",
            "board_heading",
            "board_subheading",
            "board_numbered_title",
            "yilun_table",
        } and current.role != "board_body":
            current = wenyan.StyledLine(current.text, "board_body", current.meta)
        output.append(current)
        index += 1
    return mark_yilun_large_tif_lines(output)


def chunk_yilun_board_lines(lines: list[wenyan.StyledLine]) -> list[list[wenyan.StyledLine]]:
    """保持标题/表格成组，同时为长的纯正文组做保守分页。"""

    groups: list[list[wenyan.StyledLine]] = []
    current: list[wenyan.StyledLine] = []
    for line in lines:
        role = line.role or "body"
        if role == "yilun_table":
            if current:
                current.append(line)
                groups.append(current)
                current = []
            else:
                groups.append([line])
            continue
        if role == "board_heading" and current and not (
            len(current) == 1 and (current[0].role or "") == "board_main_heading"
        ):
            groups.append(current)
            current = []
        current.append(line)
    if current:
        groups.append(current)

    chunks: list[list[wenyan.StyledLine]] = []
    for group in groups:
        if any((line.role or "") == "yilun_table" for line in group):
            chunks.append(group)
            continue
        if any((line.role or "") == "board_main_heading" for line in group):
            chunks.append(group)
            continue
        # The board prose is 22pt with first-line indentation.  Eight units
        # leaves a bottom margin for the orange footer and prevents the long
        # “阅读指导” section from reaching beyond the canvas.
        chunks.extend(wenyan.chunk_styled_lines(group, max_units=8))
    return [chunk for chunk in chunks if chunk]


def safe_rechunk_yilun(
    chunks: list[list[wenyan.StyledLine]],
    *,
    max_units: int = YILUN_MAX_PAGE_UNITS,
) -> list[list[wenyan.StyledLine]]:
    """做最后一道逐行分页，避免答案/解析后处理重新合并后越过页底。

    这里只处理本书适配器已经生成的 StyledLine，不改变公共程序的
    题型识别和答案匹配规则。若一个题目的正文与答案确实跨页，答案
    仍按原顺序继续显示，但每页都保持在模板内容区内。
    """

    result: list[list[wenyan.StyledLine]] = []
    for source_chunk in chunks:
        expanded: list[wenyan.StyledLine] = []
        for line in source_chunk:
            role = line.role or "body"
            if (
                role in {"analysis", "analysis_block", "segmentation_analysis", "yuedu_analysis"}
                and "<" not in line.text
                and wenyan.text_units(plain(line.text)) > 56
            ):
                for piece in yuedu.split_text_by_units(line.text, max_units=42):
                    expanded.append(wenyan.StyledLine(piece, line.role, line.meta))
                continue
            expanded.extend(wenyan.split_oversized_styled_line(line))

        current: list[wenyan.StyledLine] = []
        used = 0
        for line in expanded:
            if (line.role or "body") == "yilun_table":
                if current:
                    result.append(current)
                    current = []
                    used = 0
                result.append([line])
                continue
            cost = wenyan.line_page_cost(line)
            if current and used + cost > max_units:
                result.append(current)
                current = []
                used = 0
            current.append(line)
            used += cost
        if current:
            result.append(current)
    return [chunk for chunk in result if chunk]


def make_yilun_piece_slide_chunks(
    body_piece: wenyan.Piece,
    answer_piece: wenyan.Piece | None,
) -> list[list[wenyan.StyledLine]]:
    body_lines = body_piece.lines[2:]
    answer_lines = answer_piece.lines[2:] if answer_piece else []
    chunks: list[list[wenyan.StyledLine]] = []
    styled = style_segment(body_piece.title, body_lines, answer_lines or None)
    if body_piece.title == "板块必备":
        return chunk_yilun_board_lines(styled)
    judgement_section = is_yilun_judgement_group(body_lines)
    page_units = (
        YILUN_JUDGEMENT_MAX_PAGE_UNITS
        if judgement_section
        else YILUN_MAX_PAGE_UNITS
    )
    current = wenyan.chunk_styled_lines(styled, max_units=page_units)
    current = sanbu.repair_orphan_sanbu_question_chunks(current)
    current = yuedu.split_questions_after_answer_content(current)
    current = yuedu.freeze_analysis_roles(current)
    current = yuedu.attach_key_ability_to_previous_answer_chunk(current)
    rechunked = safe_rechunk_yilun(current, max_units=page_units)
    if judgement_section:
        rechunked = repair_yilun_judgement_continuation_chunks(rechunked)
    chunks.extend(rechunked)
    return chunks


def add_yilun_cover_content(root, volume: str, start: int, end: int) -> None:
    """只改封面文本，不改变模板已有的品牌图形和位置。"""

    for shape in wenyan.text_shapes(root):
        text = wenyan.get_shape_text(shape).strip()
        normalized = compact(text)
        if "高考语文大一轮复习新方案" in normalized or "大一轮复习新方案" in normalized:
            wenyan.set_shape_text(shape, BOOK_TITLE)
            wenyan.style_existing_text_shape(
                shape,
                font="宋体",
                color="000000",
                bold=True,
            )
        elif normalized in {"阅读Ⅰ信息类文本阅读", "阅读I信息类文本阅读"}:
            wenyan.set_shape_text(shape, TARGET_VOLUME)
            wenyan.style_existing_text_shape(
                shape,
                font="宋体",
                color="000000",
                bold=False,
            )


def set_yilun_slide_animations(
    root,
    animations: list[wenyan.AnimationTarget],
) -> None:
    """为本测试稿恢复“点击逐条出现”的标准 PowerPoint 动画 XML。"""

    wenyan.remove_existing_timing(root)
    if not animations:
        return

    id_value = 1

    def next_id() -> int:
        nonlocal id_value
        current = id_value
        id_value += 1
        return current

    timing = wenyan.ET.Element(wenyan.qn(wenyan.P_NS, "timing"))
    tn_lst = wenyan.ET.SubElement(timing, wenyan.qn(wenyan.P_NS, "tnLst"))
    root_par = wenyan.ET.SubElement(tn_lst, wenyan.qn(wenyan.P_NS, "par"))
    root_ctn = wenyan.ET.SubElement(
        root_par,
        wenyan.qn(wenyan.P_NS, "cTn"),
        {"id": str(next_id()), "dur": "indefinite", "restart": "never", "nodeType": "tmRoot"},
    )
    root_child = wenyan.ET.SubElement(root_ctn, wenyan.qn(wenyan.P_NS, "childTnLst"))
    seq = wenyan.ET.SubElement(
        root_child,
        wenyan.qn(wenyan.P_NS, "seq"),
        {"concurrent": "1", "nextAc": "seek"},
    )
    main_ctn = wenyan.ET.SubElement(
        seq,
        wenyan.qn(wenyan.P_NS, "cTn"),
        {"id": str(next_id()), "dur": "indefinite", "nodeType": "mainSeq"},
    )
    main_child = wenyan.ET.SubElement(main_ctn, wenyan.qn(wenyan.P_NS, "childTnLst"))
    for target in animations:
        wenyan.append_click_effect(main_child, target, next_id)

    prev_cond = wenyan.ET.SubElement(seq, wenyan.qn(wenyan.P_NS, "prevCondLst"))
    cond = wenyan.ET.SubElement(prev_cond, wenyan.qn(wenyan.P_NS, "cond"), {"evt": "onPrev", "delay": "0"})
    tgt = wenyan.ET.SubElement(cond, wenyan.qn(wenyan.P_NS, "tgtEl"))
    wenyan.ET.SubElement(tgt, wenyan.qn(wenyan.P_NS, "sldTgt"))
    next_cond = wenyan.ET.SubElement(seq, wenyan.qn(wenyan.P_NS, "nextCondLst"))
    cond = wenyan.ET.SubElement(next_cond, wenyan.qn(wenyan.P_NS, "cond"), {"evt": "onNext", "delay": "0"})
    tgt = wenyan.ET.SubElement(cond, wenyan.qn(wenyan.P_NS, "tgtEl"))
    wenyan.ET.SubElement(tgt, wenyan.qn(wenyan.P_NS, "sldTgt"))

    bld_lst = wenyan.ET.SubElement(timing, wenyan.qn(wenyan.P_NS, "bldLst"))
    for target in animations:
        attrs = {"spid": str(target.shape_id), "grpId": "0"}
        if target.text_target:
            attrs["build"] = "allAtOnce"
        else:
            attrs["animBg"] = "1"
        wenyan.ET.SubElement(bld_lst, wenyan.qn(wenyan.P_NS, "bldP"), attrs)

    ext_lst = root.find(wenyan.qn(wenyan.P_NS, "extLst"))
    if ext_lst is not None:
        root.insert(list(root).index(ext_lst), timing)
    else:
        root.append(timing)


def append_yilun_link_textbox(
    root,
    rels_root,
    *,
    x: float,
    y: float,
    w: float,
    h: float,
    name: str,
    paragraphs: list[tuple[str, list[tuple[str, str, float, str, bool]]]],
    slide_number: int | None = None,
    no_wrap: bool = False,
) -> int:
    shape_id = wenyan.next_shape_id(root)
    sp = wenyan.ET.Element(wenyan.qn(wenyan.P_NS, "sp"))

    nv_sp_pr = wenyan.ET.SubElement(sp, wenyan.qn(wenyan.P_NS, "nvSpPr"))
    c_nv_pr = wenyan.ET.SubElement(
        nv_sp_pr,
        wenyan.qn(wenyan.P_NS, "cNvPr"),
        {"id": str(shape_id), "name": f"{name} {shape_id}"},
    )
    if slide_number is not None:
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
        {
            "wrap": "none" if no_wrap else "square",
            "anchor": "t",
            "lIns": "0",
            "tIns": "0",
            "rIns": "0",
            "bIns": "0",
        },
    )
    wenyan.ET.SubElement(body_pr, wenyan.qn(wenyan.A_NS, "spAutoFit"))
    wenyan.ET.SubElement(tx_body, wenyan.qn(wenyan.A_NS, "lstStyle"))
    for align, runs in paragraphs:
        p = wenyan.ET.SubElement(tx_body, wenyan.qn(wenyan.A_NS, "p"))
        p_pr = wenyan.ET.SubElement(p, wenyan.qn(wenyan.A_NS, "pPr"), {"algn": align})
        line_spacing = wenyan.ET.SubElement(p_pr, wenyan.qn(wenyan.A_NS, "lnSpc"))
        wenyan.ET.SubElement(line_spacing, wenyan.qn(wenyan.A_NS, "spcPct"), {"val": "120000"})
        last_size = 2000
        for text, font, size, color, bold in runs:
            wenyan.append_run(p, text, font=font, size=size, color=color, bold=bold)
            last_size = int(size * 100)
        wenyan.ET.SubElement(p, wenyan.qn(wenyan.A_NS, "endParaRPr"), {"lang": "zh-CN", "sz": str(last_size)})
    wenyan.append_to_sp_tree(root, sp)
    return shape_id


def append_yilun_table_line(root, *, x: float, y: float, w: float, h: float, name: str) -> None:
    wenyan.append_rect(root, x=x, y=y, w=w, h=h, fill_color="000000", line_color=None, name=name)


def append_yilun_table_paragraphs(tx_body, text: str, *, role: str) -> None:
    ET.SubElement(tx_body, wenyan.qn(wenyan.A_NS, "lstStyle"))
    paragraphs = text.split("\n") if text else [""]
    if not paragraphs:
        paragraphs = [""]
    font, size, color, bold = wenyan.role_style(role)
    for raw_paragraph in paragraphs:
        paragraph = ET.SubElement(tx_body, wenyan.qn(wenyan.A_NS, "p"))
        p_pr = ET.SubElement(paragraph, wenyan.qn(wenyan.A_NS, "pPr"), wenyan.paragraph_properties(role))
        wenyan.add_paragraph_line_spacing(p_pr)
        if raw_paragraph:
            for inline in wenyan.parse_inline_runs(raw_paragraph):
                if not inline.text:
                    continue
                run_font = inline.font or ("宋体" if inline.answer_fill else font)
                run_bold = inline.bold if inline.bold is not None else (True if inline.answer_fill else bold)
                run_color = "FF0000" if inline.answer_fill else (inline.color or color)
                run_size = size
                run_italic = False
                run_highlight = inline.highlight
                if inline.answer_fill and role != "long_answer" and wenyan.text_units(wenyan.strip_inline_tags(inline.text)) > 18.0:
                    run_size = max(18.0, size * 0.75)
                if inline.emphasis_dot and run_font in {"楷体", "KaiTi"}:
                    run_font = "KaiTi"
                if inline.emphasis_dot:
                    wenyan.append_emphasis_dot_runs(
                        paragraph,
                        inline.text,
                        font=run_font,
                        size=run_size,
                        color=run_color,
                        bold=run_bold,
                        underline=inline.underline,
                        superscript=inline.superscript,
                        italic=run_italic,
                        underline_color=inline.underline_color,
                        highlight=run_highlight,
                    )
                else:
                    wenyan.append_run(
                        paragraph,
                        inline.text,
                        font=run_font,
                        size=run_size,
                        color=run_color,
                        bold=run_bold,
                        underline=inline.underline,
                        superscript=inline.superscript,
                        italic=run_italic,
                        underline_color=inline.underline_color,
                        highlight=run_highlight,
                    )
        ET.SubElement(paragraph, wenyan.qn(wenyan.A_NS, "endParaRPr"), {"lang": "zh-CN", "sz": str(int(size * 100))})


def append_yilun_table_cell_border(tc_pr, tag: str, *, color: str = "000000", width: int = 9525) -> None:
    line = ET.SubElement(tc_pr, wenyan.qn(wenyan.A_NS, tag), {"w": str(width)})
    solid = ET.SubElement(line, wenyan.qn(wenyan.A_NS, "solidFill"))
    ET.SubElement(solid, wenyan.qn(wenyan.A_NS, "srgbClr"), {"val": color})
    ET.SubElement(line, wenyan.qn(wenyan.A_NS, "prstDash"), {"val": "solid"})


def append_yilun_native_table(
    root,
    rels_root,
    *,
    media_assets: dict[str, bytes],
    tif_search_dirs: list[Path],
    content_box: tuple[float, float, float, float],
    table_line: wenyan.StyledLine,
    name: str,
) -> None:
    data = table_line.meta or {}
    headers = [str(item) for item in data.get("headers", [])]
    rows = [[str(cell) for cell in row] for row in data.get("rows", [])]
    if not headers or not rows:
        return

    base_x, base_y, base_w, base_h = content_box
    left = base_x + 0.18
    top = base_y + 0.08
    width = base_w - 0.36
    max_h = base_h - 0.16
    weights = yilun_table_column_weights(headers, rows)
    if len(weights) != len(headers):
        weights = [1.0 / len(headers)] * len(headers)
    col_widths = [width * weight for weight in weights]

    dense = len(rows) >= 5
    header_h = 0.30 if dense else 0.36
    row_heights = [yilun_table_row_height(row, weights, width) for row in rows]
    table_h = header_h + sum(row_heights)
    if table_h > max_h:
        overflow = table_h - max_h
        if overflow > 0.01 and row_heights:
            row_heights[-1] = max(0.28, row_heights[-1] - overflow)
            table_h = header_h + sum(row_heights)

    shape_id = wenyan.next_shape_id(root)
    gf = ET.Element(wenyan.qn(wenyan.P_NS, "graphicFrame"))
    nv = ET.SubElement(gf, wenyan.qn(wenyan.P_NS, "nvGraphicFramePr"))
    ET.SubElement(nv, wenyan.qn(wenyan.P_NS, "cNvPr"), {"id": str(shape_id), "name": f"{name} {shape_id}"})
    c_nv = ET.SubElement(nv, wenyan.qn(wenyan.P_NS, "cNvGraphicFramePr"))
    ET.SubElement(c_nv, wenyan.qn(wenyan.A_NS, "graphicFrameLocks"), {"noGrp": "1"})
    ET.SubElement(nv, wenyan.qn(wenyan.P_NS, "nvPr"))

    xfrm = ET.SubElement(gf, wenyan.qn(wenyan.P_NS, "xfrm"))
    ET.SubElement(xfrm, wenyan.qn(wenyan.A_NS, "off"), {"x": str(wenyan.inch(left)), "y": str(wenyan.inch(top))})
    ET.SubElement(xfrm, wenyan.qn(wenyan.A_NS, "ext"), {"cx": str(wenyan.inch(width)), "cy": str(wenyan.inch(table_h))})

    graphic = ET.SubElement(gf, wenyan.qn(wenyan.A_NS, "graphic"))
    graphic_data = ET.SubElement(
        graphic,
        wenyan.qn(wenyan.A_NS, "graphicData"),
        {"uri": "http://schemas.openxmlformats.org/drawingml/2006/table"},
    )
    tbl = ET.SubElement(graphic_data, wenyan.qn(wenyan.A_NS, "tbl"))
    tbl_pr = ET.SubElement(tbl, wenyan.qn(wenyan.A_NS, "tblPr"), {"firstRow": "1", "bandRow": "0", "bandCol": "0"})
    ET.SubElement(tbl_pr, wenyan.qn(wenyan.A_NS, "tableStyleId")).text = "{5940675A-B579-460E-94D1-54222C63F5DA}"

    grid = ET.SubElement(tbl, wenyan.qn(wenyan.A_NS, "tblGrid"))
    for col_w in col_widths:
        ET.SubElement(grid, wenyan.qn(wenyan.A_NS, "gridCol"), {"w": str(wenyan.inch(col_w))})

    header_role = "yilun_table_header_small" if dense else "yilun_table_header"
    cell_role = "yilun_table_cell_small" if dense else "yilun_table_cell"

    rows_data = [headers] + rows
    heights = [header_h] + row_heights
    tif_overlay_specs: list[tuple[str, float, float, float, float, str]] = []
    for row_index, (row_values, row_h) in enumerate(zip(rows_data, heights)):
        tr = ET.SubElement(tbl, wenyan.qn(wenyan.A_NS, "tr"), {"h": str(wenyan.inch(row_h))})
        normalized = row_values + [""] * (len(headers) - len(row_values))
        for col_index, (cell_text, col_w) in enumerate(zip(normalized[: len(headers)], col_widths)):
            tc = ET.SubElement(tr, wenyan.qn(wenyan.A_NS, "tc"))
            tx_body = ET.SubElement(tc, wenyan.qn(wenyan.A_NS, "txBody"))
            ET.SubElement(
                tx_body,
                wenyan.qn(wenyan.A_NS, "bodyPr"),
                {
                    "wrap": "square",
                    "rtlCol": "0",
                    "anchor": "t",
                },
            )
            append_yilun_table_paragraphs(tx_body, cell_text, role=header_role if row_index == 0 else cell_role)

            tc_pr = ET.SubElement(
                tc,
                wenyan.qn(wenyan.A_NS, "tcPr"),
                {
                    "marL": str(wenyan.inch(0.04)),
                    "marR": str(wenyan.inch(0.04)),
                    "marT": str(wenyan.inch(0.025)),
                    "marB": str(wenyan.inch(0.025)),
                },
            )
            fill = ET.SubElement(tc_pr, wenyan.qn(wenyan.A_NS, "solidFill"))
            ET.SubElement(fill, wenyan.qn(wenyan.A_NS, "srgbClr"), {"val": "FFFFFF"})
            for border_tag in ("lnL", "lnR", "lnT", "lnB"):
                append_yilun_table_cell_border(tc_pr, border_tag)

            if cell_text.strip():
                tif_overlay_specs.append(
                    (
                        cell_text,
                        left + sum(col_widths[:col_index]) + 0.05,
                        top + sum(heights[:row_index]) + 0.03,
                        max(0.08, col_w - 0.10),
                        row_h,
                        header_role if row_index == 0 else cell_role,
                    )
                )

    wenyan.append_to_sp_tree(root, gf)
    for cell_text, cell_x, cell_y, cell_w, _cell_h, cell_role_for_tif in tif_overlay_specs:
        cell_line = wenyan.StyledLine(cell_text, cell_role_for_tif)
        wenyan.append_tif_overlays(
            root,
            rels_root,
            media_assets,
            x=cell_x,
            y=cell_y,
            w=cell_w,
            lines=[cell_line],
            tif_search_dirs=tif_search_dirs,
        )


def append_yilun_table_cell(
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
) -> None:
    line = wenyan.StyledLine(text, role)
    shape_id = wenyan.append_textbox(
        root,
        x=x,
        y=y,
        w=w,
        h=h,
        lines=[line],
        name=name,
        answer_mode="normal",
    )
    # Table cells must stay inside the drawn grid.  The shared text helper
    # uses spAutoFit for normal prose, which would expand a cell past its row.
    for sp in root.findall(f".//{{{wenyan.P_NS}}}sp"):
        c_nv_pr = sp.find(f".//{{{wenyan.P_NS}}}cNvPr")
        if c_nv_pr is None or c_nv_pr.attrib.get("id") != str(shape_id):
            continue
        body_pr = sp.find(f".//{{{wenyan.A_NS}}}bodyPr")
        if body_pr is None:
            break
        for fit in list(body_pr):
            if fit.tag in {
                wenyan.qn(wenyan.A_NS, "spAutoFit"),
                wenyan.qn(wenyan.A_NS, "normAutofit"),
                wenyan.qn(wenyan.A_NS, "noAutofit"),
            }:
                body_pr.remove(fit)
        wenyan.ET.SubElement(body_pr, wenyan.qn(wenyan.A_NS, "noAutofit"))
        break
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


def append_yilun_table(
    root,
    rels_root,
    *,
    media_assets: dict[str, bytes],
    tif_search_dirs: list[Path],
    content_box: tuple[float, float, float, float],
    table_line: wenyan.StyledLine,
    name: str,
) -> None:
    append_yilun_native_table(
        root,
        rels_root,
        media_assets=media_assets,
        tif_search_dirs=tif_search_dirs,
        content_box=content_box,
        table_line=table_line,
        name=name,
    )


def yilun_visual_line_count(line: wenyan.StyledLine, width: float) -> int:
    role = line.role or "body"
    if role == "yilun_large_tif_line":
        return 1
    capacity = wenyan.line_capacity_for_width(width)
    inline_units = 0.0
    max_visual_line = 0
    for inline in wenyan.parse_inline_runs(line.text):
        if inline.tif_name:
            inline_units += 1.0
            max_visual_line = max(max_visual_line, int(inline_units // capacity))
            continue
        inline_text = re.sub(r"<[^>]+>", "", inline.text)
        inline_units += wenyan.text_units(inline_text)
        max_visual_line = max(max_visual_line, int(inline_units // capacity))
    return max(1, max_visual_line + 1)


def yilun_tif_aspect(tif_path: Path) -> float:
    try:
        from PIL import Image

        with Image.open(tif_path) as image:
            width, height = image.size
        return width / height if height else 1.0
    except Exception:
        return 1.0


def append_yilun_scaled_tif_picture(
    root,
    rels_root,
    media_assets: dict[str, bytes],
    *,
    tif_path: Path,
    x: float,
    y: float,
    max_w: float,
    max_h: float,
    name: str,
) -> int:
    package_path = wenyan.tif_package_path(tif_path)
    if package_path not in media_assets:
        media_assets[package_path] = wenyan.image_to_png_bytes(tif_path)
    rel_id = wenyan.add_image_relationship(rels_root, posixpath.relpath(package_path, "ppt/slides"))

    aspect = yilun_tif_aspect(tif_path)
    max_w *= YILUN_TIF_SCALE
    max_h *= YILUN_TIF_SCALE
    img_w = max_w
    img_h = img_w / aspect if aspect else max_h
    if img_h > max_h:
        img_h = max_h
        img_w = img_h * aspect
    img_w = min(max_w, max(0.1, img_w))
    img_h = min(max_h, max(0.1, img_h))
    return wenyan.append_picture(
        root,
        rel_id=rel_id,
        x=x + max(0.0, (max_w - img_w) / 2),
        y=y,
        w=img_w,
        h=img_h,
        name=name,
    )


def append_yilun_large_tif_line(
    root,
    rels_root,
    media_assets: dict[str, bytes],
    *,
    x: float,
    y: float,
    w: float,
    tif_names: list[str],
    tif_search_dirs: list[Path],
) -> None:
    resolved: list[tuple[str, Path, float]] = []
    for tif_name in tif_names:
        tif_path = wenyan.resolve_tif_path(tif_name, tif_search_dirs)
        if tif_path is None:
            continue
        resolved.append((tif_name, tif_path, yilun_tif_aspect(tif_path)))
    if not resolved:
        return

    left = x + 0.18
    max_w = max(1.0, w - 0.36)
    gap = 0.18
    count = len(resolved)
    slot_w = (max_w - gap * (count - 1)) / count if count > 1 else max_w
    for index, (tif_name, tif_path, aspect) in enumerate(resolved):
        max_h = YILUN_LARGE_TIF_PORTRAIT_MAX_HEIGHT if aspect < 0.9 else YILUN_LARGE_TIF_MAX_HEIGHT
        append_yilun_scaled_tif_picture(
            root,
            rels_root,
            media_assets,
            tif_path=tif_path,
            x=left + index * (slot_w + gap),
            y=y + 0.04,
            max_w=slot_w,
            max_h=max_h,
            name=f"yilun large tif {tif_name}",
        )


def append_yilun_tif_overlays(
    old_append_tif_overlays,
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
    cursor_y = 0.0
    for line in lines:
        role = line.role or "body"
        if role == "yilun_large_tif_line":
            tif_names = [
                str(name)
                for name in ((line.meta or {}).get("large_tifs") or yilun_tif_names(line.text))
                if is_yilun_large_tif_name(str(name))
            ]
            append_yilun_large_tif_line(
                root,
                rels_root,
                media_assets,
                x=x,
                y=y + cursor_y,
                w=w,
                tif_names=tif_names,
                tif_search_dirs=tif_search_dirs,
            )
        elif "data-fbd-tif" in line.text:
            cleaned, _large_markup, _large_names = strip_yilun_large_tif_tags(line.text)
            if "data-fbd-tif" in cleaned:
                old_append_tif_overlays(
                    root,
                    rels_root,
                    media_assets,
                    x=x,
                    y=y + cursor_y,
                    w=w,
                    lines=[wenyan.StyledLine(cleaned, role, line.meta)],
                    tif_search_dirs=tif_search_dirs,
                )
        cursor_y += wenyan.line_height_for_role(role) * yilun_visual_line_count(line, w)


def add_yilun_judgement_content(
    root,
    rels_root,
    piece: wenyan.Piece,
    lines: list[wenyan.StyledLine],
    *,
    media_assets: dict[str, bytes],
    tif_search_dirs: list[Path],
    content_box: tuple[float, float, float, float],
) -> None:
    """判断正误题专用：题干整体成框，答案按字段整体动画浮现。"""

    wenyan.clear_text_shapes(root, keep_substrings=("能力",))
    base_x, base_y, base_w, base_h = content_box
    y = base_y
    animations: list[wenyan.AnimationTarget] = []
    cursor_y = 0.0
    capacity = max(18.0, wenyan.line_capacity_for_width(base_w) - 5.0)
    char_w = base_w / capacity
    indent = 0.70
    text_buffer: list[wenyan.StyledLine] = []
    answer_buffer: list[wenyan.StyledLine] = []

    def line_box_height(line: wenyan.StyledLine) -> float:
        role = line.role or "body"
        visible_text = wenyan.strip_inline_tags(line.text)
        visual_units = wenyan.text_units(visible_text)
        if role == "yilun_judgement_answer":
            answer_capacity = max(18.0, capacity - 2.0)
            visual_lines = max(1, math.ceil(visual_units / answer_capacity))
        else:
            visual_lines = max(1, math.ceil(visual_units / capacity))
        line_h = wenyan.line_height_for_role(role)
        return max(
            0.34,
            visual_lines * line_h * (1.16 if role == "yilun_judgement_answer" else 1.28),
        )

    def flush_text_buffer() -> None:
        nonlocal cursor_y, text_buffer
        if not text_buffer:
            return
        box_h = sum(line_box_height(line) for line in text_buffer)
        wenyan.append_textbox(
            root,
            x=base_x,
            y=y + cursor_y,
            w=base_w,
            h=box_h,
            lines=text_buffer,
            name=f"{piece.title} judgement text",
            answer_mode="normal",
        )
        wenyan.append_tif_overlays(
            root,
            rels_root,
            media_assets,
            x=base_x,
            y=y + cursor_y,
            w=base_w,
            lines=text_buffer,
            tif_search_dirs=tif_search_dirs,
        )
        cursor_y += box_h
        text_buffer = []

    def flush_answer_buffer() -> None:
        nonlocal cursor_y, answer_buffer
        if not answer_buffer:
            return
        box_h = sum(line_box_height(line) for line in answer_buffer)
        answers = [
            str((line.meta or {}).get("judgement_answer", "")).strip()
            for line in answer_buffer
            if str((line.meta or {}).get("judgement_answer", "")).strip()
        ]
        if answers:
            first_meta = answer_buffer[0].meta or {}
            prefix_units = float(first_meta.get("judgement_prefix_units", 0.0))
            is_inline = bool(first_meta.get("judgement_inline")) and len(answers) == 1 and wenyan.text_units(answers[0]) <= 8
            if is_inline:
                answer_x = base_x + indent + prefix_units * char_w + 0.03
                answer_y = max(0.0, y + cursor_y - 0.02)
                answer_w = max(0.45, min(base_w - (answer_x - base_x) - 0.10, wenyan.estimate_answer_overlay_width_pt(answers) / 72.0 + 0.10))
            else:
                answer_x = base_x + indent
                answer_y = y + cursor_y + 0.04
                answer_w = max(1.0, base_w - indent - 0.10)
            answer_h = max(0.30, box_h + (0.01 if is_inline else 0.03))
            shape_id = append_yilun_link_textbox(
                root,
                rels_root,
                x=answer_x,
                y=answer_y,
                w=answer_w,
                h=answer_h,
                name=f"{piece.title} judgement reveal",
                paragraphs=[("l", [(answer, "宋体", 24, "FF0000", False)]) for answer in answers],
                no_wrap=is_inline,
            )
            animations.append(wenyan.AnimationTarget(shape_id=shape_id, effect="appear", text_target=False))
        cursor_y += box_h
        answer_buffer = []

    for line in lines:
        if (line.role or "body") == "yilun_judgement_answer":
            flush_text_buffer()
            answer_buffer.append(line)
            continue
        flush_answer_buffer()
        text_buffer.append(line)
    flush_text_buffer()
    flush_answer_buffer()
    wenyan.set_slide_animations(root, animations)


def add_yilun_content(
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
    table_index = next(
        (index for index, line in enumerate(chunk) if (line.role or "") == "yilun_table"),
        None,
    )
    if any((line.role or "") == "yilun_judgement_answer" for line in chunk):
        add_yilun_judgement_content(
            root,
            rels_root,
            piece,
            chunk,
            media_assets=media_assets,
            tif_search_dirs=tif_search_dirs,
            content_box=content_box,
        )
        return
    if table_index is not None:
        wenyan.clear_text_shapes(root, keep_substrings=("能力",))
        base_x, base_y, base_w, base_h = content_box
        prefix = chunk[:table_index]
        table_y = base_y
        table_h = base_h
        if prefix:
            prefix_h = min(
                1.45,
                sum(wenyan.line_height_for_role(line.role or "body") for line in prefix)
                + 0.06 * max(0, len(prefix) - 1),
            )
            wenyan.append_textbox(
                root,
                x=base_x,
                y=base_y,
                w=base_w,
                h=prefix_h,
                lines=prefix,
                name=f"{piece.title} table heading",
                answer_mode="normal",
            )
            table_y = base_y + prefix_h + 0.08
            table_h = max(1.0, base_h - (table_y - base_y))
        append_yilun_table(
            root,
            rels_root,
            media_assets=media_assets,
            tif_search_dirs=tif_search_dirs,
            content_box=(base_x, table_y, base_w, table_h),
            table_line=chunk[table_index],
            name=f"{piece.title} table",
        )
        wenyan.set_slide_animations(root, [])
        return
    fallback_add_content(
        root,
        rels_root,
        piece,
        chunk,
        media_assets=media_assets,
        tif_search_dirs=tif_search_dirs,
        content_box=content_box,
    )


def yilun_group_label(volume: str) -> str:
    if volume.startswith("专题一"):
        return "专题一"
    if volume.startswith("专题二"):
        return "专题二"
    return volume


def yilun_top_directory_groups(pieces: list[wenyan.Piece]) -> list[tuple[str, str, list[wenyan.Piece]]]:
    result: list[tuple[str, str, list[wenyan.Piece]]] = []
    seen: set[str] = set()
    for piece in pieces:
        label = yilun_group_label(piece.volume)
        if label in seen:
            continue
        grouped = [candidate for candidate in pieces if yilun_group_label(candidate.volume) == label]
        result.append((label, piece.volume, grouped))
        seen.add(label)
    return result


def yilun_detail_directory_groups(pieces: list[wenyan.Piece]) -> list[tuple[str, str, list[wenyan.Piece]]]:
    result: list[tuple[str, str, list[wenyan.Piece]]] = []
    for label, volume, grouped in yilun_top_directory_groups(pieces):
        detail_pieces = [piece for piece in grouped if piece.title.startswith("突破")]
        if label.startswith("专题") and detail_pieces:
            result.append((label, volume, detail_pieces))
    return result


def add_yilun_top_directory_content(
    root,
    rels_root,
    groups: list[tuple[str, str, list[wenyan.Piece]]],
    targets: dict[str, int],
) -> None:
    for sp in wenyan.text_shapes(root):
        text = wenyan.get_shape_text(sp).strip()
        if text and "课件目录" not in text:
            wenyan.set_shape_text(sp, "")

    y_positions = [1.34, 2.48, 3.62, 4.76]
    for y, (label, _volume, _grouped) in zip(y_positions, groups):
        slide_number = targets.get(label)
        append_yilun_link_textbox(
            root,
            rels_root,
            x=4.35,
            y=y,
            w=5.90,
            h=0.58,
            name=f"一轮一级目录 {label}",
            slide_number=slide_number,
            paragraphs=[("l", [("· ", "黑体", 32, "000000", True), (label, "黑体", 32, "000000", True)])],
        )


def add_yilun_group_directory_content(
    root,
    rels_root,
    group_label: str,
    group_volume: str,
    group_pieces: list[wenyan.Piece],
    title_slide_numbers: dict[int, int],
) -> None:
    for sp in wenyan.text_shapes(root):
        text = wenyan.get_shape_text(sp).strip()
        if text and "各篇目点击即可直达" not in text and "课件目录" not in text:
            wenyan.set_shape_text(sp, "")

    header_text = f"「{group_label}」" if group_volume == group_label else f"「{group_label} {group_volume[len(group_label):].strip()}」"
    append_yilun_link_textbox(
        root,
        rels_root,
        x=1.00,
        y=0.55,
        w=11.40,
        h=0.65,
        name=f"一轮二级目录标题 {group_label}",
        paragraphs=[("ctr", [(header_text, "黑体", 40, "E77818", False)])],
        no_wrap=True,
    )

    columns = 1 if len(group_pieces) <= 5 else 2
    split_at = math.ceil(len(group_pieces) / columns)
    left_x = 3.35
    right_x = 7.75
    start_y = 1.72
    step = 0.78 if len(group_pieces) > 6 else 0.92
    for index, piece in enumerate(group_pieces):
        column = 0 if index < split_at else 1
        row = index if column == 0 else index - split_at
        x = left_x if column == 0 else right_x
        y = start_y + row * step
        append_yilun_link_textbox(
            root,
            rels_root,
            x=x,
            y=y,
            w=4.20 if columns == 2 else 8.4,
            h=0.52,
            name=f"一轮二级目录 {piece.num}",
            slide_number=title_slide_numbers[piece.num],
            paragraphs=[("l", [("· ", "黑体", 28, "000000", True), (piece.title, "黑体", 28, "000000", True)])],
        )


def build_yilun_deck(
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
        top_groups = yilun_top_directory_groups(pieces)
        detail_groups = yilun_detail_directory_groups(pieces)
        directory_slide_count = 1 + len(detail_groups)

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

        detail_slide_numbers = {
            volume_key: 4 + index for index, (_label, volume_key, _group_pieces) in enumerate(detail_groups)
        }
        top_targets: dict[str, int] = {}
        for label, volume_key, grouped in top_groups:
            top_targets[label] = detail_slide_numbers.get(volume_key, title_slide_numbers[grouped[0].num])

        def add_from(src: int, edit_func=lambda root, rels_root: None) -> None:
            xml, rels = wenyan.make_slide_parts(zin, src, edit_func)
            generated_slides.append((src, xml, rels))

        add_from(1, lambda root, rels_root: add_yilun_cover_content(root, volume, start, end))
        add_from(2)
        add_from(3, lambda root, rels_root: add_yilun_top_directory_content(root, rels_root, top_groups, top_targets))
        for group_label, group_volume, group_pieces in detail_groups:
            add_from(
                3,
                lambda root, rels_root, group_label=group_label, group_volume=group_volume, group_pieces=group_pieces: add_yilun_group_directory_content(
                    root,
                    rels_root,
                    group_label,
                    group_volume,
                    group_pieces,
                    title_slide_numbers,
                ),
            )

        for piece, chunks in piece_chunks:
            add_from(title_src, lambda root, rels_root, piece=piece: wenyan.add_piece_title(root, piece))
            for chunk in chunks:
                add_from(
                    content_src,
                    lambda root, rels_root, piece=piece, chunk=chunk: wenyan.add_content(
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
    wenyan.postprocess_tif_positions_with_powerpoint(
        output,
        multiline_answer_min_width_pt=multiline_answer_min_width_pt,
    )


def install_yilun_hooks() -> tuple[object, ...]:
    """安装本书的切片 hook，退出时完整恢复公共模块状态。"""

    old_hooks = yuedu.install_yuedu_hooks()
    old_yilun_extract_answer_items = wenyan.extract_answer_items
    old_yilun_line_has_fillable_blank = wenyan.line_has_fillable_blank
    old_yilun_fill_lines_with_answer_values = wenyan.fill_lines_with_answer_values
    old_yilun_strip_fillable_blanks_from_line = wenyan.strip_fillable_blanks_from_line
    old_yilun_line_page_cost = wenyan.line_page_cost
    old_yuedu_split_text_by_units = yuedu.split_text_by_units
    old_yuedu_centered_long_answer_lines = yuedu.yuedu_centered_long_answer_lines
    old_yuedu_is_open_answer_group = yuedu.is_open_answer_group
    old_yilun_role_style = wenyan.role_style
    old_yilun_paragraph_properties = wenyan.paragraph_properties
    old_yilun_add_content = wenyan.add_content
    old_yilun_append_answer_overlays = wenyan.append_answer_overlays
    old_yilun_append_tif_overlays = wenyan.append_tif_overlays

    def yilun_line_page_cost(line: wenyan.StyledLine) -> int:
        role = line.role or "body"
        if role == "yilun_table":
            return 99
        if role == "board_body":
            return 2
        if role == "yilun_judgement_answer":
            if line.meta and line.meta.get("judgement_inline"):
                return 0.5
            return 1
        cost = old_yilun_line_page_cost(line)
        if line.meta and line.meta.get("yilun_judgement_text"):
            visible = wenyan.strip_inline_tags(line.text)
            return max(cost, int(math.ceil(wenyan.text_units(visible) / 24.0)))
        if role == "long_answer":
            return max(YILUN_LONG_ANSWER_PAGE_COST, cost)
        if role == "yilun_judgement_answer":
            visible = wenyan.strip_inline_tags(line.text)
            if line.meta and line.meta.get("judgement_inline"):
                return max(1, int(math.ceil(wenyan.text_units(visible) / 28.0)))
            return max(1, int(math.ceil(wenyan.text_units(visible) / 22.0)))
        return cost

    def yilun_role_style(role: str) -> tuple[str, float, str, bool]:
        if role == "board_main_heading":
            return "黑体", 24, "000000", True
        if role == "board_heading":
            return "黑体", 24, "000000", True
        if role == "board_subheading":
            return "黑体", 24, "000000", True
        if role == "board_numbered_title":
            return "黑体", 24, "000000", True
        if role == "board_body":
            return "宋体", 24, "000000", False
        if role == "yilun_table_header":
            return "黑体", 14, "000000", True
        if role == "yilun_table_header_small":
            return "黑体", 14, "000000", True
        if role == "yilun_table_cell":
            return "宋体", 14, "000000", False
        if role == "yilun_table_cell_small":
            return "宋体", 14, "000000", False
        if role == "yilun_judgement_answer":
            return "宋体", 24, "FF0000", False
        if role in {"yilun_large_tif_line", "yilun_large_tif_spacer"}:
            return "宋体", 24, "000000", False
        return old_yilun_role_style(role)

    def yilun_paragraph_properties(role: str) -> dict[str, str]:
        if role == "yilun_judgement_answer":
            return old_yilun_paragraph_properties("body")
        if role in {
            "board_main_heading",
            "board_heading",
            "board_subheading",
            "board_numbered_title",
            "board_body",
            "yilun_table_header",
            "yilun_table_header_small",
            "yilun_table_cell",
            "yilun_table_cell_small",
            "yilun_large_tif_line",
            "yilun_large_tif_spacer",
        }:
            attrs = old_yilun_paragraph_properties("body")
            attrs["indent"] = "0"
            attrs["marL"] = "0"
            attrs["algn"] = "ctr" if "table_header" in role else "l"
            return attrs
        if role == "board_body":
            attrs = old_yilun_paragraph_properties("body")
            attrs["indent"] = str(wenyan.FIRST_LINE_INDENT_EMU)
            attrs["marL"] = "0"
            attrs["algn"] = "just"
            return attrs
        return old_yilun_paragraph_properties(role)

    def yilun_split_text_by_units(text: str, max_units: float = 42.0) -> list[str]:
        # The shared answer animator uses 30-unit fragments for very long
        # analysis paragraphs.  This book's 24pt body makes those fragments
        # create one extra visual row at the bottom of a few slides; a wider
        # fragment still fits the content box and keeps the paragraph intact.
        if max_units == 30:
            max_units = 36.0
        return old_yuedu_split_text_by_units(text, max_units=max_units)

    def yilun_is_open_answer_group(lines: list[str]) -> bool:
        if is_yilun_judgement_group(lines):
            return False
        return old_yuedu_is_open_answer_group(lines)

    def yilun_centered_long_answer_lines(
        answer: str,
        *,
        max_units: float = yuedu.YUEDU_LONG_ANSWER_TEXT_UNITS,
    ) -> list[wenyan.StyledLine]:
        # For this book the reference deck keeps long answers as one visible
        # paragraph and lets PowerPoint wrap naturally.  Returning a single
        # underlined run avoids the artificial mid-sentence breaks that the
        # generic splitter introduces.
        text = re.sub(r"\s+", "", plain(answer)).strip()
        if not text:
            return [wenyan.StyledLine("", "long_answer")]
        return [wenyan.StyledLine(yuedu.yuedu_fixed_answer_line(text), "long_answer")]

    def yilun_add_content(
        root,
        rels_root,
        piece: wenyan.Piece,
        chunk: list[wenyan.StyledLine],
        *,
        media_assets: dict[str, bytes],
        tif_search_dirs: list[Path],
        content_box: tuple[float, float, float, float],
    ) -> None:
        add_yilun_content(
            root,
            rels_root,
            piece,
            chunk,
            media_assets=media_assets,
            tif_search_dirs=tif_search_dirs,
            content_box=content_box,
            fallback_add_content=old_yilun_add_content,
        )

    def yilun_append_answer_overlays(
        root,
        *,
        x: float,
        y: float,
        w: float,
        h: float,
        lines: list[wenyan.StyledLine],
        name: str,
    ) -> list[wenyan.AnimationTarget]:
        analysis_roles = {"analysis", "analysis_block", "segmentation_analysis", "yuedu_analysis"}
        if not any((line.role or "") in analysis_roles for line in lines):
            return old_yilun_append_answer_overlays(root, x=x, y=y, w=w, h=h, lines=lines, name=name)
        if any((line.role or "") == "long_answer" for line in lines):
            return old_yilun_append_answer_overlays(root, x=x, y=y, w=w, h=h, lines=lines, name=name)

        animations: list[wenyan.AnimationTarget] = []
        event_index = 0
        cursor_y = 0.0
        capacity = max(18.0, wenyan.line_capacity_for_width(w) - 2.0)

        def visual_line_count(line: wenyan.StyledLine) -> int:
            role = line.role or "body"
            if role in analysis_roles:
                return max(1, math.ceil(wenyan.text_units(plain(line.text)) / capacity))
            return max(1, wenyan.estimate_line_count(line.text))

        def overlay_line_step(line: wenyan.StyledLine) -> float:
            return visual_line_count(line) * wenyan.line_height_for_role(line.role or "body") * 1.20

        def append_event(event_lines: list[wenyan.StyledLine], event_y: float, index: int) -> None:
            if not any(wenyan.strip_inline_tags(line.text).strip() for line in event_lines):
                return
            event_h = max(0.28, sum(overlay_line_step(line) for line in event_lines) + 0.04)
            shape_id = wenyan.append_textbox(
                root,
                x=x,
                y=event_y,
                w=w,
                h=min(h, event_h),
                lines=event_lines,
                name=f"{name} answer {index + 1}",
                answer_mode="normal",
            )
            animations.append(wenyan.AnimationTarget(shape_id=shape_id, effect="appear", text_target=False))

        index = 0
        while index < len(lines):
            line = lines[index]
            role = line.role or "body"
            if role in analysis_roles:
                group: list[wenyan.StyledLine] = []
                start_y = cursor_y
                while index < len(lines) and (lines[index].role or "") in analysis_roles:
                    group.append(lines[index])
                    cursor_y += overlay_line_step(lines[index])
                    index += 1
                append_event(group, y + start_y, event_index)
                event_index += 1
                continue

            if role in wenyan.ANIMATED_LINE_ROLES:
                append_event([line], y + cursor_y, event_index)
                event_index += 1
                cursor_y += overlay_line_step(line)
                index += 1
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
            cursor_y += overlay_line_step(line)
            index += 1
        return animations

    def yilun_extract_answer_items(answer_lines: list[str] | None) -> dict[int, wenyan.AnswerItem]:
        items = old_yilun_extract_answer_items(answer_lines)
        return normalize_yilun_two_part_judgement_answers(items)

    def yilun_line_has_fillable_blank(line: str) -> bool:
        if old_yilun_line_has_fillable_blank(line):
            return True
        return bool(re.search(r'<span\s+class="blank"[^>]*>\s*</span>', line))

    def yilun_fill_lines_with_answer_values(
        lines: list[str],
        answers: list[str],
    ) -> tuple[list[str], int]:
        filled_lines, used = old_yilun_fill_lines_with_answer_values(lines, answers)
        if used >= len(answers):
            return filled_lines, used

        index = used
        blank_pattern = re.compile(r'<span\s+class="blank"[^>]*>\s*</span>')
        result: list[str] = []
        for line in filled_lines:
            def repl(match: re.Match[str]) -> str:
                nonlocal index
                if index >= len(answers):
                    return match.group(0)
                answer = answers[index]
                index += 1
                return wenyan.answer_span(answer, underline=True, pad_underline=True)

            result.append(blank_pattern.sub(repl, line))
        return result, index

    def yilun_strip_fillable_blanks_from_line(line: str) -> str:
        line = old_yilun_strip_fillable_blanks_from_line(line)
        line = re.sub(r'<span\s+class="blank"[^>]*>\s*</span>', "", line)
        return line.strip()

    def yilun_append_tif_overlays(
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
        append_yilun_tif_overlays(
            old_yilun_append_tif_overlays,
            root,
            rels_root,
            media_assets,
            x=x,
            y=y,
            w=w,
            lines=lines,
            tif_search_dirs=tif_search_dirs,
        )

    old_make = wenyan.make_piece_slide_chunks
    old_cover = wenyan.add_cover_content
    old_set_animations = wenyan.set_slide_animations
    wenyan.make_piece_slide_chunks = make_yilun_piece_slide_chunks
    wenyan.add_cover_content = add_yilun_cover_content
    wenyan.set_slide_animations = set_yilun_slide_animations
    wenyan.extract_answer_items = yilun_extract_answer_items
    wenyan.line_has_fillable_blank = yilun_line_has_fillable_blank
    wenyan.fill_lines_with_answer_values = yilun_fill_lines_with_answer_values
    wenyan.strip_fillable_blanks_from_line = yilun_strip_fillable_blanks_from_line
    wenyan.line_page_cost = yilun_line_page_cost
    wenyan.role_style = yilun_role_style
    wenyan.paragraph_properties = yilun_paragraph_properties
    wenyan.add_content = yilun_add_content
    wenyan.append_answer_overlays = yilun_append_answer_overlays
    wenyan.append_tif_overlays = yilun_append_tif_overlays
    yuedu.split_text_by_units = yilun_split_text_by_units
    yuedu.yuedu_centered_long_answer_lines = yilun_centered_long_answer_lines
    yuedu.is_open_answer_group = yilun_is_open_answer_group
    return old_hooks + (
        old_make,
        old_cover,
        old_set_animations,
        old_yilun_extract_answer_items,
        old_yilun_line_has_fillable_blank,
        old_yilun_fill_lines_with_answer_values,
        old_yilun_strip_fillable_blanks_from_line,
        old_yilun_line_page_cost,
        old_yuedu_split_text_by_units,
        old_yuedu_centered_long_answer_lines,
        old_yuedu_is_open_answer_group,
        old_yilun_role_style,
        old_yilun_paragraph_properties,
        old_yilun_add_content,
        old_yilun_append_answer_overlays,
        old_yilun_append_tif_overlays,
    )


def restore_yilun_hooks(old_hooks: tuple[object, ...]) -> None:
    old_yuedu_hooks = old_hooks[:-16]
    (
        old_make,
        old_cover,
        old_set_animations,
        old_extract_answer_items,
        old_line_has_fillable_blank,
        old_fill_lines_with_answer_values,
        old_strip_fillable_blanks_from_line,
        old_line_page_cost,
        old_split_text_by_units,
        old_centered_long_answer_lines,
        old_is_open_answer_group,
        old_role_style,
        old_paragraph_properties,
        old_add_content,
        old_append_answer_overlays,
        old_append_tif_overlays,
    ) = old_hooks[-16:]
    wenyan.make_piece_slide_chunks = old_make
    wenyan.add_cover_content = old_cover
    wenyan.set_slide_animations = old_set_animations
    wenyan.extract_answer_items = old_extract_answer_items
    wenyan.line_has_fillable_blank = old_line_has_fillable_blank
    wenyan.fill_lines_with_answer_values = old_fill_lines_with_answer_values
    wenyan.strip_fillable_blanks_from_line = old_strip_fillable_blanks_from_line
    wenyan.line_page_cost = old_line_page_cost
    wenyan.role_style = old_role_style
    wenyan.paragraph_properties = old_paragraph_properties
    wenyan.add_content = old_add_content
    wenyan.append_answer_overlays = old_append_answer_overlays
    wenyan.append_tif_overlays = old_append_tif_overlays
    yuedu.split_text_by_units = old_split_text_by_units
    yuedu.yuedu_centered_long_answer_lines = old_centered_long_answer_lines
    yuedu.is_open_answer_group = old_is_open_answer_group
    yuedu.restore_yuedu_hooks(old_yuedu_hooks)


@isolated_generation
def generate_yilun_pptx(
    *,
    md_run: Path,
    template: Path = DEFAULT_TEMPLATE,
    output_dir: Path = DEFAULT_OUTPUT_DIR,
    progress: Callable[[str], None] | None = None,
    group_indexes: list[int] | None = None,
    group_ranges: list[dict[str, object]] | None = None,
) -> list[Path]:
    def emit(message: str) -> None:
        (progress or print)(message)

    body_md, answer_md = find_md_files(md_run)
    board_cache = build_yilun_board_cache(body_md, answer_md)
    tif_dirs = wenyan.infer_tif_search_dirs(md_run)
    if DEFAULT_SOURCE_DIR.exists() and DEFAULT_SOURCE_DIR not in tif_dirs:
        tif_dirs.append(DEFAULT_SOURCE_DIR)
    global YILUN_DYNAMIC_TABLE_SPECS
    source_body_fbd = DEFAULT_BODY_FBD if DEFAULT_BODY_FBD.exists() else None
    manifest_path = md_run / "conversion_manifest.json"
    if manifest_path.exists():
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            for item in manifest.get("outputs", []):
                if item.get("role") == "正文" and item.get("source"):
                    candidate = Path(str(item["source"]))
                    if candidate.exists():
                        source_body_fbd = candidate
                    break
        except Exception:
            pass
    YILUN_DYNAMIC_TABLE_SPECS = (
        extract_yilun_dynamic_table_specs(source_body_fbd)
        if source_body_fbd is not None
        else {}
    )
    output_dir.mkdir(parents=True, exist_ok=True)

    emit(f"正文 MD：{body_md}")
    emit(f"答案 MD：{answer_md}")
    emit(f"板块数：{len(board_cache)}")
    emit(f"源 FBD 表格种类：{len(YILUN_DYNAMIC_TABLE_SPECS)}")
    emit(f"输出目录：{output_dir}")

    specs = YILUN_DECK_SPECS
    if group_indexes:
        specs = [YILUN_DECK_SPECS[index - 1] for index in group_indexes if 1 <= index <= len(YILUN_DECK_SPECS)]
    if group_ranges is not None:
        specs = []
        for index, group in enumerate(group_ranges, 1):
            board = str(group["volume"])
            start = int(group["start"])
            end = int(group["end"])
            keys = {tuple(key) for key in group["unit_keys"]}  # type: ignore[arg-type]
            selected = [piece for piece in board_cache.get(board, ([], {}))[0] if (piece.volume, piece.num) in keys]
            specs.append({"filename": f"{index}{board} 第{start}-{end}篇.pptx", "board": board, "piece_start": start, "piece_end": end, "cover_start": start, "cover_end": end, "selected_override": selected})

    outputs: list[Path] = []
    old_hooks = install_yilun_hooks()
    try:
        for spec in specs:
            board_key = str(spec["board"])
            board_pieces, answer_map = board_cache.get(board_key, ([], {}))
            selected = [
                piece
                for piece in board_pieces
                if int(spec["piece_start"]) <= piece.num <= int(spec["piece_end"])
            ]
            selected = list(spec.get("selected_override") or selected)
            if not selected:
                emit(f"跳过，无篇章：{spec['filename']}")
                continue
            output = output_dir / str(spec["filename"])
            emit(f"生成：{output.name}，板块：{board_key}，片段：{len(selected)}")
            build_yilun_deck(
                template,
                output,
                selected,
                answer_map,
                board_key,
                int(spec["cover_start"]),
                int(spec["cover_end"]),
                tif_search_dirs=tif_dirs,
            )
            outputs.append(output)
    finally:
        restore_yilun_hooks(old_hooks)

    manifest = {
        "md_run": str(md_run),
        "template": str(template),
        "output_dir": str(output_dir),
        "generated": [str(path) for path in outputs],
    }
    (output_dir / "yilun_test_generation_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
        newline="\n",
    )
    return outputs


def scan_yilun_group_units(md_run: Path) -> list[dict[str, object]]:
    """Return detected board sections and numbered units for dynamic grouping."""

    body_md, answer_md = find_md_files(md_run)
    cache = build_yilun_board_cache(body_md, answer_md)
    units: list[dict[str, object]] = []
    for board, (pieces, _answers) in cache.items():
        units.extend({"volume": board, "num": piece.num, "title": piece.title} for piece in pieces)
    return units


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="《高考语文大一轮复习新方案》测试生成器")
    parser.add_argument("--md-run", default="", help="已有 MD run 目录；省略时先转换 FBD")
    parser.add_argument("--body-fbd", default=str(DEFAULT_BODY_FBD))
    parser.add_argument("--answer-fbd", default=str(DEFAULT_ANSWER_FBD))
    parser.add_argument("--md-output-root", default=str(DEFAULT_MD_ROOT))
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE))
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR))
    parser.add_argument("--group-index", type=int, action="append", default=[], help="只生成指定序号的 pptx，可重复")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    md_run = Path(args.md_run) if args.md_run else convert_yilun_fbd_to_md(
        body_fbd=Path(args.body_fbd),
        answer_fbd=Path(args.answer_fbd),
        output_root=Path(args.md_output_root),
    )
    generate_yilun_pptx(
        md_run=md_run,
        template=Path(args.template),
        output_dir=Path(args.output_dir),
        group_indexes=args.group_index or None,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

#!/usr/bin/env python3
"""Convert Founder/FBD source text into Markdown with semantic inline styles.

The converter keeps text content plus the layout semantics needed by later
pipeline stages: dotted emphasis, straight underline, wavy underline, blanks,
and superscript note markers.
"""

from __future__ import annotations

import argparse
import html
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Callable


STYLE_BLOCK = """<style>
body {
  font-family: "Microsoft YaHei", "SimSun", serif;
  line-height: 1.85;
  font-size: 16px;
}
.dot {
  text-emphasis: filled dot;
  text-emphasis-position: under;
  -webkit-text-emphasis: filled dot;
  -webkit-text-emphasis-position: under;
}
.uline {
  text-decoration-line: underline;
  text-underline-offset: 0.18em;
  text-decoration-thickness: 1px;
}
.wave {
  text-decoration-line: underline;
  text-decoration-style: wavy;
  text-underline-offset: 0.18em;
  text-decoration-thickness: 1px;
}
.bold {
  font-weight: 700;
}
.blank {
  display: inline-block;
  border-bottom: 1px solid currentColor;
  height: 0.9em;
  vertical-align: -0.08em;
}
.fbd-tif-char {
  height: 1em;
  width: auto;
  vertical-align: -0.12em;
}
sup {
  font-size: 0.72em;
  line-height: 0;
  vertical-align: super;
}
</style>
"""


CONTENT_START_MARKERS = (
    "基础篇30篇",
    "基础篇",
    "编写概述",
    "答案解析",
)


@dataclass(frozen=True)
class ConvertStats:
    chars: int
    lines: int
    dots: int
    underlines: int
    waves: int
    blanks: int
    superscripts: int
    tif_kept: int = 0
    tif_ignored: int = 0
    tif_unknown: int = 0
    bolds: int = 0


TifResolver = Callable[[str, str], str | None]


def choose_input_file() -> Path:
    """Open a standard file picker when the script is run without arguments."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception as exc:  # pragma: no cover - depends on local desktop.
        raise SystemExit(
            "未传入 .fbd 文件路径，且无法打开文件选择框。"
            "请改用：python tools/fbd_to_md.py path/to/file.fbd"
        ) from exc

    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(
        title="选择 FBD 文件",
        filetypes=[("FBD files", "*.fbd"), ("All files", "*.*")],
    )
    root.destroy()
    if not path:
        raise SystemExit("未选择文件，已退出。")
    return Path(path)


def read_fbd(path: Path, encoding: str) -> str:
    try:
        return path.read_text(encoding=encoding)
    except UnicodeDecodeError:
        if encoding.lower().replace("-", "") != "gb18030":
            return path.read_text(encoding="gb18030")
        raise


def normalize_tif_key(name: str) -> str:
    return Path(name.strip()).name.lower()


def tif_markdown_tag(name: str) -> str:
    safe_name = html.escape(Path(name).name, quote=True)
    return (
        f'<img class="fbd-tif-char" '
        f'src="{safe_name}" alt="{safe_name}" data-fbd-tif="{safe_name}" />'
    )


def load_tif_map(path: Path | None) -> tuple[dict[str, str], set[str]]:
    if not path or not path.exists():
        return {}, set()
    data = json.loads(path.read_text(encoding="utf-8"))
    keep_names = set(data.get("keep") or [])
    # Legacy compatibility: old maps used "replace" to store hand-entered
    # characters. Under the new workflow those same files are preserved as tif.
    keep_names.update((data.get("replace") or {}).keys())
    replace = {normalize_tif_key(name): tif_markdown_tag(name) for name in keep_names}
    ignore = {normalize_tif_key(name) for name in (data.get("ignore") or [])}
    return replace, ignore


def save_tif_map(path: Path, replace: dict[str, str], ignore: set[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    data = {
        "keep": sorted(replace),
        "ignore": sorted(ignore),
    }
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8", newline="\n")


def tif_name_from_command(command_body: str) -> str | None:
    body = command_body.strip()
    upper = body.upper()
    for prefix in ("XC", "TP"):
        if upper.startswith(prefix):
            body = body[len(prefix) :]
            break
    match = re.search(r"([^,;，；\s〗]+?\.(?:tif|tiff))", body, flags=re.I)
    if not match:
        return None
    return Path(match.group(1)).name


def replace_tif_commands(
    text: str,
    *,
    tif_replacements: dict[str, str] | None = None,
    tif_ignores: set[str] | None = None,
    tif_resolver: TifResolver | None = None,
) -> tuple[str, int, int, int]:
    replacements = tif_replacements if tif_replacements is not None else {}
    ignores = tif_ignores if tif_ignores is not None else set()
    kept = 0
    ignored = 0
    unknown = 0

    def repl(match: re.Match[str]) -> str:
        nonlocal kept, ignored, unknown
        command_body = match.group(1)
        tif_name = tif_name_from_command(command_body)
        if not tif_name:
            unknown += 1
            return ""

        key = normalize_tif_key(tif_name)
        if key in replacements:
            kept += 1
            return replacements[key]
        if key in ignores:
            ignored += 1
            return ""
        if tif_resolver:
            value = tif_resolver(tif_name, match.group(0))
            if value:
                kept += 1
                return value
            ignored += 1
            return ""

        unknown += 1
        return ""

    text = re.sub(r"〖([^〗]*?\.(?:tif|tiff)[^〗]*)〗", repl, text, flags=re.I)
    return text, kept, ignored, unknown


def find_auto_start(text: str) -> int:
    """Skip the initial page template block when a known content marker exists."""
    positions = [text.find(marker) for marker in CONTENT_START_MARKERS]
    positions = [pos for pos in positions if pos >= 0]
    return min(positions) if positions else 0


def apply_single_char_dots(text: str) -> str:
    """Handle commands such as 〖ZZ1〗宜 and 〖ZZ2〗左右."""
    out: list[str] = []
    i = 0
    dot_remaining = 0
    marker_1 = "〖ZZ1〗"
    marker_2 = "〖ZZ2〗"

    while i < len(text):
        if text.startswith(marker_1, i):
            dot_remaining = 1
            i += len(marker_1)
            continue
        if text.startswith(marker_2, i):
            dot_remaining = 2
            i += len(marker_2)
            continue

        ch = text[i]
        if dot_remaining > 0 and ch not in "\n\r\t ":
            out.append(f'<span class="dot">{ch}</span>')
            dot_remaining -= 1
        else:
            out.append(ch)
        i += 1

    return "".join(out)


def replace_range_styles(text: str) -> str:
    """Handle 〖ZZ(Q〗...〖ZZ)〗, 〖ZZ(Z〗...〖ZZ)〗, and 〖ZZ(〗...〖ZZ)〗."""

    def span(class_name: str):
        def repl(match: re.Match[str]) -> str:
            return f'<span class="{class_name}">{match.group(1)}</span>'

        return repl

    text = re.sub(
        r"〖ZZ\(Q(?:,[^〗]*)?〗(?:〖=〗)?(.*?)〖ZZ\)〗",
        span("wave"),
        text,
        flags=re.S,
    )
    text = re.sub(
        r"〖ZZ\(Z(?:,[^〗]*)?〗(?:〖=〗)?(.*?)〖ZZ\)〗",
        span("uline"),
        text,
        flags=re.S,
    )
    text = re.sub(r"〖ZZ\(〗(.*?)〖ZZ\)〗", span("dot"), text, flags=re.S)
    return text


def replace_bold_styles(text: str) -> str:
    """Keep FBD emphasis ranges such as 〖=Y(〗...〖=〗 as bold text."""

    def repl(match: re.Match[str]) -> str:
        return f'<span class="bold">{match.group(1)}</span>'

    return re.sub(r"〖=Y\(〗(.*?)〖=〗", repl, text, flags=re.S)


def replace_blanks(text: str) -> str:
    """Convert FBD fill commands and placeholder runs into visible blanks."""

    def cd_repl(match: re.Match[str]) -> str:
        count = int(match.group(1))
        width = max(3.0, count * 0.55)
        return f'<span class="blank" style="width:{width:.1f}em"></span>'

    def placeholder_repl(match: re.Match[str]) -> str:
        width = max(2.0, len(match.group(0)) * 0.9)
        return f'<span class="blank" style="width:{width:.1f}em"></span>'

    text = re.sub(r"〖CD#(\d+)〗", cd_repl, text)
    text = re.sub(r"〓+", placeholder_repl, text)
    return text


def replace_superscripts(text: str) -> str:
    """Convert the FBD superscript marker U+E00B into HTML sup tags."""
    text = re.sub("\ue00b(【注】|[①-⑳])", r"<sup>\1</sup>", text)
    return text.replace("\ue00b", "")


def normalize_private_controls(text: str, field_break: str) -> str:
    """Normalize private-use control characters used by this FBD dialect."""
    field_replacement = {
        "newline": "\n",
        "space": " ",
        "remove": "",
    }[field_break]

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\ue003", "\n")
    text = text.replace("\ue004", "\n")
    text = text.replace("\ue005", field_replacement)
    text = text.replace("\ue008", "")
    text = text.replace("\ue009", "")
    return text


def clean_nonsemantic_commands(text: str) -> str:
    """Remove FBD commands after semantic commands have been converted."""
    text = re.sub(r"〖[^〗]*〗", "", text)
    text = re.sub(r"\ue000[^\ue002\n]{0,16}\ue002", "", text)
    text = text.replace("\ue5e6", " ")
    text = re.sub(r"[\ue000-\uf8ff]", "", text)
    return text


def tidy_lines(text: str) -> str:
    text = text.replace("\t", " ")
    text = re.sub(r"[ \u3000]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)

    lines: list[str] = []
    for raw_line in text.split("\n"):
        line = raw_line.strip()
        if not line:
            if lines and lines[-1] != "":
                lines.append("")
            continue
        if re.fullmatch(r"[\s,，.。;；:：、\-—~～·|｜]+", line):
            continue
        lines.append(line)

    while lines and lines[-1] == "":
        lines.pop()
    return "\n".join(lines) + "\n"


def convert_fbd_to_markdown(
    source_text: str,
    *,
    auto_start: bool = True,
    start_marker: str | None = None,
    include_css: bool = True,
    field_break: str = "newline",
    tif_replacements: dict[str, str] | None = None,
    tif_ignores: set[str] | None = None,
    tif_resolver: TifResolver | None = None,
) -> tuple[str, ConvertStats]:
    if start_marker:
        start = source_text.find(start_marker)
        if start >= 0:
            source_text = source_text[start:]
    elif auto_start:
        source_text = source_text[find_auto_start(source_text) :]

    text = html.escape(source_text, quote=False)
    text, tif_kept, tif_ignored, tif_unknown = replace_tif_commands(
        text,
        tif_replacements=tif_replacements,
        tif_ignores=tif_ignores,
        tif_resolver=tif_resolver,
    )
    text = normalize_private_controls(text, field_break=field_break)
    text = replace_superscripts(text)
    text = apply_single_char_dots(text)
    text = replace_bold_styles(text)
    text = replace_range_styles(text)
    text = replace_blanks(text)
    text = clean_nonsemantic_commands(text)
    text = tidy_lines(text)

    markdown = f"{STYLE_BLOCK}\n{text}" if include_css else text
    stats = ConvertStats(
        chars=len(markdown),
        lines=markdown.count("\n"),
        dots=markdown.count('class="dot"'),
        underlines=markdown.count('class="uline"'),
        waves=markdown.count('class="wave"'),
        blanks=markdown.count('class="blank"'),
        superscripts=markdown.count("<sup>"),
        tif_kept=tif_kept,
        tif_ignored=tif_ignored,
        tif_unknown=tif_unknown,
        bolds=markdown.count('class="bold"'),
    )
    return markdown, stats


def default_output_path(input_path: Path, output_dir: Path | None) -> Path:
    target_dir = output_dir if output_dir else input_path.parent / "output" / "md"
    return target_dir / f"{input_path.stem}_格式保留.md"


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="将 .fbd 文件转换成保留加点、下划线、波浪线、填空线和上标信息的 Markdown。",
    )
    parser.add_argument("input", nargs="?", help="待转换的 .fbd 文件；省略时弹出文件选择框。")
    parser.add_argument("-o", "--output", help="输出 .md 文件路径。")
    parser.add_argument("--output-dir", help="输出目录；未指定时使用 输入目录/output/md。")
    parser.add_argument("--encoding", default="gb18030", help="FBD 文本编码，默认 gb18030。")
    parser.add_argument("--start-marker", help="从指定文字第一次出现处开始转换。")
    parser.add_argument(
        "--no-auto-start",
        action="store_true",
        help="不要自动跳过开头版式模板，转换整个 FBD 文本。",
    )
    parser.add_argument("--no-css", action="store_true", help="不写入 Markdown 内置 CSS。")
    parser.add_argument(
        "--tif-map",
        help="tif 替换映射 JSON；未指定时自动尝试使用 FBD 同目录下的 fbd_tif_map.json。",
    )
    parser.add_argument(
        "--field-break",
        choices=("newline", "space", "remove"),
        default="newline",
        help="U+E005 字段分隔符的处理方式，默认 newline。",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    input_path = Path(args.input) if args.input else choose_input_file()
    if not input_path.exists():
        raise SystemExit(f"找不到输入文件：{input_path}")

    output_dir = Path(args.output_dir) if args.output_dir else None
    output_path = Path(args.output) if args.output else default_output_path(input_path, output_dir)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    raw = read_fbd(input_path, args.encoding)
    tif_map_path = Path(args.tif_map) if args.tif_map else input_path.parent / "fbd_tif_map.json"
    tif_replacements, tif_ignores = load_tif_map(tif_map_path)
    markdown, stats = convert_fbd_to_markdown(
        raw,
        auto_start=not args.no_auto_start,
        start_marker=args.start_marker,
        include_css=not args.no_css,
        field_break=args.field_break,
        tif_replacements=tif_replacements,
        tif_ignores=tif_ignores,
    )
    output_path.write_text(markdown, encoding="utf-8", newline="\n")

    print(f"输入：{input_path}")
    print(f"输出：{output_path}")
    print(f"字符数：{stats.chars}")
    print(f"行数：{stats.lines}")
    print(
        "样式统计："
        f"加点 {stats.dots}，"
        f"横线 {stats.underlines}，"
        f"波浪线 {stats.waves}，"
        f"加粗 {stats.bolds}，"
        f"填空线 {stats.blanks}，"
        f"上标 {stats.superscripts}，"
        f"tif保留 {stats.tif_kept}，"
        f"tif忽略 {stats.tif_ignored}，"
        f"tif未知 {stats.tif_unknown}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

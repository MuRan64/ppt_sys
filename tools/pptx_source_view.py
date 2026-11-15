#!/usr/bin/env python3
"""Inspect a PPTX file as Open XML source.

PPTX files are zip packages. This script opens a deck, reads the package
structure, extracts slide text/shape metadata, and optionally dumps all XML
parts into a folder so the file can be reviewed like source code.
"""

from __future__ import annotations

import argparse
import posixpath
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.dom import minidom
from xml.etree import ElementTree as ET


DEFAULT_PPTX = Path(
    r"E:\ppt项目书籍\语文\语文两本新书\文言文三步进阶模板及说明\文言文三步进阶模板.pptx"
)

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


@dataclass(frozen=True)
class PackagePart:
    name: str
    size: int
    compressed_size: int


@dataclass(frozen=True)
class ShapeInfo:
    kind: str
    shape_id: str
    name: str
    placeholder: str
    text: str


@dataclass(frozen=True)
class SlideInfo:
    index: int
    part_name: str
    relationship_id: str
    text: str
    shapes: list[ShapeInfo]


def local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag


def xml_namespace(tag: str) -> str:
    if tag.startswith("{") and "}" in tag:
        return tag[1:].split("}", 1)[0]
    return ""


def parse_xml(data: bytes) -> ET.Element:
    return ET.fromstring(data)


def pretty_xml(data: bytes) -> str:
    try:
        parsed = minidom.parseString(data)
        pretty = parsed.toprettyxml(indent="  ", encoding="utf-8").decode("utf-8")
        return "\n".join(line for line in pretty.splitlines() if line.strip()) + "\n"
    except Exception:
        return data.decode("utf-8", errors="replace")


def markdown_escape(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", "<br>")


def read_zip_xml(zf: zipfile.ZipFile, part_name: str) -> ET.Element | None:
    try:
        return parse_xml(zf.read(part_name))
    except KeyError:
        return None


def list_parts(zf: zipfile.ZipFile) -> list[PackagePart]:
    parts = []
    for info in zf.infolist():
        if info.is_dir():
            continue
        parts.append(PackagePart(info.filename, info.file_size, info.compress_size))
    return sorted(parts, key=lambda part: part.name)


def read_relationships(zf: zipfile.ZipFile, rels_part: str) -> dict[str, str]:
    root = read_zip_xml(zf, rels_part)
    if root is None:
        return {}

    relationships: dict[str, str] = {}
    for rel in root:
        if local_name(rel.tag) != "Relationship":
            continue
        rel_id = rel.attrib.get("Id", "")
        target = rel.attrib.get("Target", "")
        if rel_id and target:
            relationships[rel_id] = target
    return relationships


def resolve_relationship_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base = posixpath.dirname(source_part)
    return posixpath.normpath(posixpath.join(base, target))


def numeric_slide_key(name: str) -> tuple[int, str]:
    match = re.search(r"slide(\d+)\.xml$", name)
    return (int(match.group(1)) if match else 999999, name)


def ordered_slide_parts(zf: zipfile.ZipFile) -> list[tuple[str, str]]:
    presentation = read_zip_xml(zf, "ppt/presentation.xml")
    rels = read_relationships(zf, "ppt/_rels/presentation.xml.rels")
    ordered: list[tuple[str, str]] = []

    if presentation is not None:
        slide_id_list = presentation.find("p:sldIdLst", NS)
        if slide_id_list is not None:
            for slide_id in slide_id_list.findall("p:sldId", NS):
                rid = slide_id.attrib.get(f"{{{NS['r']}}}id", "")
                target = rels.get(rid)
                if target:
                    ordered.append((rid, resolve_relationship_target("ppt/presentation.xml", target)))

    if ordered:
        return ordered

    fallback = [
        ("", part)
        for part in zf.namelist()
        if part.startswith("ppt/slides/slide") and part.endswith(".xml")
    ]
    return sorted(fallback, key=lambda pair: numeric_slide_key(pair[1]))


def collect_text(node: ET.Element) -> str:
    texts: list[str] = []
    for child in node.iter():
        if local_name(child.tag) == "t" and xml_namespace(child.tag) == NS["a"]:
            if child.text:
                texts.append(child.text)
    return "".join(texts).strip()


def first_descendant(node: ET.Element, tag_name: str) -> ET.Element | None:
    for child in node.iter():
        if local_name(child.tag) == tag_name:
            return child
    return None


def shape_kind(node: ET.Element) -> str:
    return local_name(node.tag)


def shape_info(node: ET.Element) -> ShapeInfo:
    c_nv_pr = first_descendant(node, "cNvPr")
    shape_id = c_nv_pr.attrib.get("id", "") if c_nv_pr is not None else ""
    name = c_nv_pr.attrib.get("name", "") if c_nv_pr is not None else ""

    placeholder = ""
    ph = first_descendant(node, "ph")
    if ph is not None:
        placeholder = ph.attrib.get("type", "") or "placeholder"

    return ShapeInfo(
        kind=shape_kind(node),
        shape_id=shape_id,
        name=name,
        placeholder=placeholder,
        text=collect_text(node),
    )


def parse_slide(zf: zipfile.ZipFile, index: int, rid: str, part_name: str) -> SlideInfo:
    root = read_zip_xml(zf, part_name)
    if root is None:
        return SlideInfo(index, part_name, rid, "", [])

    shape_tags = {"sp", "pic", "graphicFrame", "cxnSp", "grpSp"}
    shapes = [
        shape_info(node)
        for node in root.iter()
        if local_name(node.tag) in shape_tags and xml_namespace(node.tag) == NS["p"]
    ]
    return SlideInfo(
        index=index,
        part_name=part_name,
        relationship_id=rid,
        text=collect_text(root),
        shapes=shapes,
    )


def content_types_summary(zf: zipfile.ZipFile) -> list[tuple[str, str, str]]:
    root = read_zip_xml(zf, "[Content_Types].xml")
    if root is None:
        return []

    rows: list[tuple[str, str, str]] = []
    for child in root:
        name = local_name(child.tag)
        if name == "Default":
            rows.append(("Default", child.attrib.get("Extension", ""), child.attrib.get("ContentType", "")))
        elif name == "Override":
            rows.append(("Override", child.attrib.get("PartName", ""), child.attrib.get("ContentType", "")))
    return rows


def dump_xml_parts(zf: zipfile.ZipFile, output_dir: Path) -> int:
    xml_root = output_dir / "xml"
    count = 0
    for name in zf.namelist():
        if not (name.endswith(".xml") or name.endswith(".rels")):
            continue
        target = xml_root / name
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(pretty_xml(zf.read(name)), encoding="utf-8", newline="\n")
        count += 1
    return count


def media_parts(zf: zipfile.ZipFile) -> list[PackagePart]:
    return [
        PackagePart(info.filename, info.file_size, info.compress_size)
        for info in zf.infolist()
        if info.filename.startswith("ppt/media/") and not info.is_dir()
    ]


def xml_excerpt(zf: zipfile.ZipFile, part_name: str, max_chars: int) -> str:
    try:
        text = pretty_xml(zf.read(part_name))
    except KeyError:
        return ""
    if len(text) <= max_chars:
        return text
    return text[:max_chars] + "\n<!-- ... truncated in summary; see xml folder for full source ... -->\n"


def write_archive_tree(parts: list[PackagePart], output_dir: Path) -> Path:
    tree_path = output_dir / "archive_tree.txt"
    lines = ["# PPTX archive tree", ""]
    for part in parts:
        lines.append(f"{part.name}\t{part.size} bytes\tcompressed {part.compressed_size} bytes")
    tree_path.write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")
    return tree_path


def write_markdown_report(
    pptx_path: Path,
    output_dir: Path,
    parts: list[PackagePart],
    slides: list[SlideInfo],
    content_types: list[tuple[str, str, str]],
    media: list[PackagePart],
    zf: zipfile.ZipFile,
    max_xml_chars: int,
) -> Path:
    report_path = output_dir / f"{pptx_path.stem}_source_view.md"
    lines: list[str] = []

    lines.append(f"# {pptx_path.name} 源码查看")
    lines.append("")
    lines.append("PPTX 是一个 ZIP 包，内部主体是 Office Open XML。下面是包结构、每页文本/形状摘要，以及关键 XML 片段。完整 XML 已导出到同级 `xml/` 文件夹。")
    lines.append("")
    lines.append(f"- 源文件：`{pptx_path}`")
    lines.append(f"- 文件大小：{pptx_path.stat().st_size} bytes")
    lines.append(f"- 包内文件数：{len(parts)}")
    lines.append(f"- 幻灯片数：{len(slides)}")
    lines.append(f"- 媒体文件数：{len(media)}")
    lines.append("")

    lines.append("## 幻灯片顺序与文本")
    lines.append("")
    lines.append("| 页码 | XML part | r:id | 文本预览 |")
    lines.append("| --- | --- | --- | --- |")
    for slide in slides:
        preview = slide.text[:180] + ("..." if len(slide.text) > 180 else "")
        lines.append(
            f"| {slide.index} | `{slide.part_name}` | `{slide.relationship_id}` | {markdown_escape(preview)} |"
        )
    lines.append("")

    lines.append("## 每页形状摘要")
    lines.append("")
    for slide in slides:
        lines.append(f"### Slide {slide.index}: `{slide.part_name}`")
        lines.append("")
        lines.append("| kind | id | name | placeholder | text |")
        lines.append("| --- | --- | --- | --- | --- |")
        for shape in slide.shapes:
            preview = shape.text[:140] + ("..." if len(shape.text) > 140 else "")
            lines.append(
                "| "
                + " | ".join(
                    [
                        f"`{shape.kind}`",
                        f"`{shape.shape_id}`",
                        markdown_escape(shape.name),
                        markdown_escape(shape.placeholder),
                        markdown_escape(preview),
                    ]
                )
                + " |"
            )
        lines.append("")

    lines.append("## Content Types")
    lines.append("")
    lines.append("| kind | key | content type |")
    lines.append("| --- | --- | --- |")
    for kind, key, content_type in content_types:
        lines.append(f"| {kind} | `{markdown_escape(key)}` | `{content_type}` |")
    lines.append("")

    lines.append("## 媒体文件")
    lines.append("")
    if media:
        lines.append("| part | size | compressed |")
        lines.append("| --- | ---: | ---: |")
        for part in media:
            lines.append(f"| `{part.name}` | {part.size} | {part.compressed_size} |")
    else:
        lines.append("无媒体文件。")
    lines.append("")

    lines.append("## 包内文件列表")
    lines.append("")
    lines.append("| part | size | compressed |")
    lines.append("| --- | ---: | ---: |")
    for part in parts:
        lines.append(f"| `{part.name}` | {part.size} | {part.compressed_size} |")
    lines.append("")

    lines.append("## 关键 XML 片段")
    lines.append("")
    key_parts = ["[Content_Types].xml", "ppt/presentation.xml"]
    key_parts.extend(slide.part_name for slide in slides[:3])
    seen: set[str] = set()
    for part_name in key_parts:
        if part_name in seen:
            continue
        seen.add(part_name)
        lines.append(f"### `{part_name}`")
        lines.append("")
        lines.append("```xml")
        lines.append(xml_excerpt(zf, part_name, max_xml_chars).rstrip())
        lines.append("```")
        lines.append("")

    report_path.write_text("\n".join(lines), encoding="utf-8", newline="\n")
    return report_path


def inspect_pptx(pptx_path: Path, output_dir: Path, dump_xml: bool, max_xml_chars: int) -> None:
    if not pptx_path.exists():
        raise FileNotFoundError(f"找不到 PPTX：{pptx_path}")
    if pptx_path.suffix.lower() != ".pptx":
        raise ValueError(f"输入文件不是 .pptx：{pptx_path}")

    output_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(pptx_path) as zf:
        parts = list_parts(zf)
        slide_pairs = ordered_slide_parts(zf)
        slides = [
            parse_slide(zf, index=i + 1, rid=rid, part_name=part_name)
            for i, (rid, part_name) in enumerate(slide_pairs)
        ]
        content_types = content_types_summary(zf)
        media = media_parts(zf)

        tree_path = write_archive_tree(parts, output_dir)
        report_path = write_markdown_report(
            pptx_path,
            output_dir,
            parts,
            slides,
            content_types,
            media,
            zf,
            max_xml_chars=max_xml_chars,
        )
        xml_count = dump_xml_parts(zf, output_dir) if dump_xml else 0

    print(f"输入 PPTX：{pptx_path}")
    print(f"输出目录：{output_dir}")
    print(f"源码报告：{report_path}")
    print(f"包文件树：{tree_path}")
    print(f"幻灯片数：{len(slides)}")
    print(f"包内文件数：{len(parts)}")
    print(f"XML 导出数：{xml_count}")


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="打开 PPTX 并导出类似源码的 Open XML 查看结果。")
    parser.add_argument(
        "pptx",
        nargs="?",
        default=str(DEFAULT_PPTX),
        help="PPTX 文件路径；省略时使用脚本内置模板路径。",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        help="输出目录；默认 output/pptx_source/<pptx 文件名>/。",
    )
    parser.add_argument(
        "--no-dump-xml",
        action="store_true",
        help="只生成总览报告，不导出完整 XML 文件夹。",
    )
    parser.add_argument(
        "--max-xml-chars",
        type=int,
        default=16000,
        help="Markdown 报告中每个 XML 片段最多展示的字符数，默认 16000。",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    pptx_path = Path(args.pptx)
    output_dir = (
        Path(args.output_dir)
        if args.output_dir
        else Path.cwd() / "output" / "pptx_source" / pptx_path.stem
    )
    inspect_pptx(
        pptx_path=pptx_path,
        output_dir=output_dir,
        dump_xml=not args.no_dump_xml,
        max_xml_chars=args.max_xml_chars,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Book-scoped template catalog shared by the GUI and book adapters."""

from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
import zipfile

from app_resources import resource_path


@dataclass(frozen=True)
class BuiltinTemplate:
    subject: str
    book: str
    slot: str
    label: str
    path: Path


def builtin_template_path(filename: str) -> Path:
    return resource_path(Path("static") / "templates" / filename)


def load_template_catalog() -> list[BuiltinTemplate]:
    catalog_path = builtin_template_path("catalog.json")
    data = json.loads(catalog_path.read_text(encoding="utf-8"))
    if not isinstance(data, dict) or not isinstance(data.get("templates"), list):
        raise ValueError("内置模板目录配置无效。")
    templates = []
    seen = set()
    for entry in data["templates"]:
        if not isinstance(entry, dict):
            raise ValueError("内置模板条目无效。")
        values = [entry.get(key) for key in ("subject", "book", "slot", "label", "filename")]
        if not all(isinstance(value, str) and value.strip() for value in values):
            raise ValueError("内置模板条目缺少科目、书名、模块、名称或文件名。")
        subject, book, slot, label, filename = values
        if Path(filename).name != filename or "/" in filename or "\\" in filename:
            raise ValueError("内置模板文件名不能包含目录。")
        key = (subject, book, slot, label)
        if key in seen:
            raise ValueError(f"内置模板名称重复：{label}")
        seen.add(key)
        templates.append(BuiltinTemplate(subject, book, slot, label, builtin_template_path(filename)))
    return templates


def templates_for_book(subject: str, book: str, slot: str = "primary") -> list[BuiltinTemplate]:
    return [template for template in load_template_catalog() if (template.subject, template.book, template.slot) == (subject, book, slot)]


def validate_pptx_template(path: Path, *, label: str = "PPTX 模板", minimum_slides: int = 6) -> None:
    if not path.is_file():
        raise FileNotFoundError(f"{label}不存在：{path}")
    if path.suffix.lower() != ".pptx" or not zipfile.is_zipfile(path):
        raise ValueError(f"{label}不是有效的 PPTX 文件：{path}")
    with zipfile.ZipFile(path) as package:
        required = {"[Content_Types].xml", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"}
        required.update(f"ppt/slides/slide{number}.xml" for number in range(1, minimum_slides + 1))
        if not required.issubset(package.namelist()):
            raise ValueError(f"{label}缺少配套模板所需的页面结构：{path}")

from __future__ import annotations

import argparse
import asyncio
import json
from pathlib import Path
from typing import Any

from pptx import Presentation


def _find_layout_by_name(prs: Presentation, keyword: str):
    for layout in prs.slide_layouts:
        if keyword in layout.name:
            return layout
    raise ValueError(f"Layout not found with keyword: {keyword}")


def _clear_all_slides(prs: Presentation) -> None:
    sld_id_lst = prs.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        rel_id = sld_id.rId
        prs.part.drop_rel(rel_id)
        sld_id_lst.remove(sld_id)


def _fill_placeholders(slide, texts: list[str]) -> None:
    placeholders = []
    for shape in slide.shapes:
        if not getattr(shape, "is_placeholder", False):
            continue
        if not getattr(shape, "has_text_frame", False):
            continue
        idx = shape.placeholder_format.idx
        placeholders.append((idx, shape))
    placeholders.sort(key=lambda item: item[0])
    for i, text in enumerate(texts):
        if i >= len(placeholders):
            break
        placeholders[i][1].text = text


def _set_shape_texts(slide, texts: list[str]) -> None:
    """Overwrite visible text boxes/placeholders with meaningful content."""
    targets = [shape for shape in slide.shapes if getattr(shape, "has_text_frame", False)]
    for i, text in enumerate(texts):
        if i >= len(targets):
            break
        targets[i].text = text


def _retain_slide_indices(prs: Presentation, keep_indices: set[int]) -> None:
    sld_id_lst = prs.slides._sldIdLst
    for idx in range(len(prs.slides) - 1, -1, -1):
        if idx in keep_indices:
            continue
        sld_id = sld_id_lst[idx]
        rel_id = sld_id.rId
        prs.part.drop_rel(rel_id)
        del sld_id_lst[idx]


def _build_examples_from_template(template_path: Path, output_path: Path, title: str) -> int:
    prs = Presentation(str(template_path))
    _clear_all_slides(prs)

    cover = prs.slides.add_slide(_find_layout_by_name(prs, "封面_Cover"))
    _fill_placeholders(cover, [title, "Reusable OpenClaw skill for company PPT templates"])

    toc = prs.slides.add_slide(_find_layout_by_name(prs, "目录页_Content"))
    _fill_placeholders(
        toc,
        [
            "目录",
            "项目目标\n统一模板能力\n示例输出\n复杂页面策略\n后续路线图",
        ],
    )

    section = prs.slides.add_slide(_find_layout_by_name(prs, "章节页_Section page"))
    _fill_placeholders(section, ["项目概览 - Project Overview", "将模板能力标准化为可迁移 Skill", "01"])

    content = prs.slides.add_slide(_find_layout_by_name(prs, "标准内容页（小标题）_Standard page with subtitle"))
    _fill_placeholders(
        content,
        [
            "本项目做了什么",
            "1) 标准化 OpenClaw skill 目录\n2) 自动提取 page catalog\n3) 支持基础/复杂两种 demo 生成模式",
            "能力摘要 - Capability Summary",
        ],
    )

    chart_layout = _find_layout_by_name(prs, "仅标题（带背景）_Title only with bg")
    chart_slide = prs.slides.add_slide(chart_layout)
    _set_shape_texts(
        chart_slide,
        [
            "复杂页面处理策略",
            "通过 page_catalog 识别图表/复杂版式页",
            "优先复用模板原型，减少视觉偏差与重绘成本",
        ],
    )

    end = prs.slides.add_slide(_find_layout_by_name(prs, "封底_End page"))
    _fill_placeholders(end, ["Thanks."])

    prs.save(str(output_path))
    return len(prs.slides)


def _pick_indices_from_catalog(catalog: dict) -> list[int]:
    picks = []
    wanted = ["cover", "toc", "section", "content", "chart_or_data", "end"]
    pages = catalog.get("pages", [])
    for page_type in wanted:
        found = next((p["index"] for p in pages if p.get("page_type_guess") == page_type), None)
        if found is not None:
            picks.append(found)
    if not picks:
        picks = [0, 1, 2, 3, 21, 55]
    return sorted(set(picks))


def _build_complex_examples_from_catalog(template_path: Path, catalog_path: Path, output_path: Path, title: str) -> int:
    if not catalog_path.exists():
        raise FileNotFoundError(f"Catalog not found: {catalog_path}")
    catalog = json.loads(catalog_path.read_text(encoding="utf-8"))
    selected = _pick_indices_from_catalog(catalog)

    prs = Presentation(str(template_path))
    _retain_slide_indices(prs, set(selected))

    narrative = {
        "cover": [title, "Complex demo generated via page catalog"],
        "toc": ["目录", "能力架构\n生成流程\n复杂页面复用\n质量校验\n发布方式"],
        "section": ["能力架构 - Architecture", "从模板语义到可执行生成策略", "01"],
        "content": ["生成流程", "Catalog 提取 -> 页面选择 -> 语义填充 -> 输出校验", "Workflow"],
        "chart_or_data": ["复杂页面示例", "该页来自 catalog 自动选择的复杂模板页", "保留视觉风格并替换关键文案"],
        "end": ["Thanks."],
        "generic": ["项目页面", "该页面由 catalog 自动选择并进行文案替换"],
    }

    page_meta = {p["index"]: p for p in catalog.get("pages", [])}
    for slide, original_idx in zip(prs.slides, selected):
        page_type = page_meta.get(original_idx, {}).get("page_type_guess", "generic")
        texts = narrative.get(page_type, narrative["generic"])
        _set_shape_texts(slide, texts)

    prs.save(str(output_path))
    return len(prs.slides)


async def handler(input: dict[str, Any], _context: Any) -> dict[str, Any]:
    skill_root = Path(__file__).resolve().parent.parent
    template_file = input.get("template_file") or "PPT_Template.pptx"
    template_path = skill_root / template_file
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")

    mode = input.get("mode", "basic")
    title = input.get("title", "Demo Deck")
    toc_items = input.get("toc_items", ["Background", "Solution", "Plan"])
    section_title = input.get("section_title", "Section - Overview")
    content_title = input.get("content_title", "Content")
    content_body = input.get("content_body", "This slide is generated with template reuse.")
    output_filename = input.get("output_filename", "openclaw_generated.pptx")
    output_path = skill_root / output_filename

    if mode == "examples":
        slide_count = _build_examples_from_template(template_path, output_path, title)
        return {
            "output_path": str(output_path.resolve()),
            "slide_count": slide_count,
            "message": "Project-intro example deck generated",
        }

    if mode == "complex":
        catalog_path = skill_root / "assets" / "page_catalog.json"
        slide_count = _build_complex_examples_from_catalog(template_path, catalog_path, output_path, title)
        return {
            "output_path": str(output_path.resolve()),
            "slide_count": slide_count,
            "message": "Complex demo generated using extracted page catalog",
        }

    prs = Presentation(str(template_path))
    _clear_all_slides(prs)

    cover = prs.slides.add_slide(_find_layout_by_name(prs, "封面_Cover"))
    _fill_placeholders(cover, [title, "Generated by OpenClaw skill"])

    toc = prs.slides.add_slide(_find_layout_by_name(prs, "目录页_Content"))
    _fill_placeholders(toc, ["目录", "\n".join(toc_items)])

    section = prs.slides.add_slide(_find_layout_by_name(prs, "章节页_Section page"))
    _fill_placeholders(section, [section_title, "Section summary", "01"])

    content = prs.slides.add_slide(_find_layout_by_name(prs, "标准内容页（小标题）_Standard page with subtitle"))
    _fill_placeholders(content, [content_title, content_body, "Subtitle"])

    end = prs.slides.add_slide(_find_layout_by_name(prs, "封底_End page"))
    _fill_placeholders(end, ["Thanks."])

    prs.save(str(output_path))

    return {
        "output_path": str(output_path.resolve()),
        "slide_count": len(prs.slides),
        "message": "PPT generated successfully",
    }


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run ppt-template-builder locally.")
    parser.add_argument("--mode", default="basic", choices=["basic", "examples", "complex"])
    parser.add_argument("--title", default="Demo Deck")
    parser.add_argument("--output", default="openclaw_generated.pptx")
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    payload = {
        "mode": args.mode,
        "title": args.title,
        "output_filename": args.output,
    }
    result = asyncio.run(handler(payload, None))
    print(result)


if __name__ == "__main__":
    main()

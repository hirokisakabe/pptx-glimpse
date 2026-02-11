#!/usr/bin/env python3
"""
LibreOffice VRT 用の PPTX フィクスチャを python-pptx で生成する。

Usage:
    python3 scripts/vrt/generate_fixtures.py
"""

import os

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "../../tests/vrt/libreoffice-fixtures")

SLIDE_WIDTH = 9144000
SLIDE_HEIGHT = 5143500


def new_presentation():
    """スライドサイズを固定した新しいプレゼンテーションを生成する。"""
    prs = Presentation()
    prs.slide_width = Emu(SLIDE_WIDTH)
    prs.slide_height = Emu(SLIDE_HEIGHT)
    return prs


def create_basic_shapes():
    """基本図形: rect, ellipse, roundRect (ソリッド塗り + 枠線)"""
    prs = new_presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

    shapes_def = [
        (MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.5), Inches(2.5), Inches(2),
         RGBColor(0x44, 0x72, 0xC4), "Rectangle"),
        (MSO_SHAPE.OVAL, Inches(3.5), Inches(0.5), Inches(2.5), Inches(2),
         RGBColor(0xED, 0x7D, 0x31), "Oval"),
        (MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.5), Inches(0.5), Inches(2.5), Inches(2),
         RGBColor(0xA5, 0xA5, 0xA5), "Rounded Rect"),
        (MSO_SHAPE.DIAMOND, Inches(0.5), Inches(3), Inches(2.5), Inches(2),
         RGBColor(0xFF, 0xC0, 0x00), "Diamond"),
        (MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(3.5), Inches(3), Inches(2.5), Inches(2),
         RGBColor(0x5B, 0x9B, 0xD5), "Triangle"),
        (MSO_SHAPE.HEXAGON, Inches(6.5), Inches(3), Inches(2.5), Inches(2),
         RGBColor(0x70, 0xAD, 0x47), "Hexagon"),
    ]

    for shape_type, left, top, width, height, color, _label in shapes_def:
        shape = slide.shapes.add_shape(shape_type, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = color
        shape.line.color.rgb = RGBColor(0x33, 0x33, 0x33)
        shape.line.width = Pt(1.5)

    prs.save(os.path.join(OUTPUT_DIR, "lo-basic-shapes.pptx"))
    print("  Created: lo-basic-shapes.pptx")


def create_text_formatting():
    """テキスト書式: 太字、イタリック、フォントサイズ、色、配置"""
    prs = new_presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    tests = [
        {"text": "Bold Text", "bold": True, "size": Pt(24)},
        {"text": "Italic Text", "italic": True, "size": Pt(24)},
        {"text": "Large 36pt", "size": Pt(36)},
        {"text": "Red Text", "size": Pt(24), "color": RGBColor(0xFF, 0x00, 0x00)},
        {"text": "Center Aligned", "size": Pt(24), "align": PP_ALIGN.CENTER},
        {"text": "Right Aligned", "size": Pt(24), "align": PP_ALIGN.RIGHT},
    ]

    for i, t in enumerate(tests):
        col = i % 2
        row = i // 2
        left = Inches(0.3 + col * 4.7)
        top = Inches(0.3 + row * 1.7)

        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, left, top, Inches(4.4), Inches(1.4)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(0xF0, 0xF0, 0xF0)
        shape.line.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
        shape.line.width = Pt(0.5)

        tf = shape.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = t["text"]

        if "align" in t:
            p.alignment = t["align"]

        run = p.runs[0]
        run.font.name = "Liberation Sans"
        run.font.size = t.get("size", Pt(18))

        if t.get("bold"):
            run.font.bold = True
        if t.get("italic"):
            run.font.italic = True
        if "color" in t:
            run.font.color.rgb = t["color"]

    prs.save(os.path.join(OUTPUT_DIR, "lo-text-formatting.pptx"))
    print("  Created: lo-text-formatting.pptx")


def create_fill_and_lines():
    """塗りと線: ソリッド塗り各色 + 枠線スタイル"""
    prs = new_presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    colors = [
        (RGBColor(0xFF, 0x63, 0x84), "Pink"),
        (RGBColor(0x36, 0xA2, 0xEB), "Blue"),
        (RGBColor(0xFF, 0xCE, 0x56), "Yellow"),
    ]

    for i, (color, _label) in enumerate(colors):
        left = Inches(0.5 + i * 3)
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(0.5),
            Inches(2.5), Inches(2)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = color
        shape.line.color.rgb = RGBColor(0x33, 0x33, 0x33)
        shape.line.width = Pt(2)

    # 枠線のみの図形（塗りなし）
    no_fill_shapes = [
        (MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(3), Inches(2.5), Inches(2),
         RGBColor(0x44, 0x72, 0xC4), Pt(1)),
        (MSO_SHAPE.OVAL, Inches(3.5), Inches(3), Inches(2.5), Inches(2),
         RGBColor(0xED, 0x7D, 0x31), Pt(2)),
        (MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.5), Inches(3), Inches(2.5), Inches(2),
         RGBColor(0x70, 0xAD, 0x47), Pt(3)),
    ]

    for shape_type, left, top, width, height, line_color, line_width in no_fill_shapes:
        shape = slide.shapes.add_shape(shape_type, left, top, width, height)
        shape.fill.background()
        shape.line.color.rgb = line_color
        shape.line.width = line_width

    prs.save(os.path.join(OUTPUT_DIR, "lo-fill-and-lines.pptx"))
    print("  Created: lo-fill-and-lines.pptx")


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print("Generating LibreOffice VRT fixtures...")
    create_basic_shapes()
    create_text_formatting()
    create_fill_and_lines()
    print("Done!")


if __name__ == "__main__":
    main()

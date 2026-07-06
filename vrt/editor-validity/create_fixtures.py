#!/usr/bin/env python3
"""
Generate PPTX fixtures for editor-validity tests with python-pptx.

The fixtures are source / expected pairs consumed by editor-validity.test.ts,
plus a basic-shapes source used by slide topology and shape add/delete checks.
basic-shapes.pptx duplicates the renderer VRT fixture of the same name on
purpose: the editor-validity suite must stay runnable without generating the
vrt/libreoffice/ fixture set.

Usage:
    python3 vrt/editor-validity/create_fixtures.py
"""

import base64
import os
import tempfile

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Emu, Inches, Pt

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "fixtures")

SLIDE_WIDTH = 9144000
SLIDE_HEIGHT = 5143500


def new_presentation():
    """Generate a new presentation with fixed slide sizes."""
    prs = Presentation()
    prs.slide_width = Emu(SLIDE_WIDTH)
    prs.slide_height = Emu(SLIDE_HEIGHT)
    return prs


def create_basic_shapes():
    """Basic shapes: 6 presets from rect to hexagon (solid fill + border)"""
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

    prs.save(os.path.join(OUTPUT_DIR, "basic-shapes.pptx"))
    print("  Created: basic-shapes.pptx")


def create_editor_validity_text_fixture(filename, text):
    """Fixture pair for editor-core text replacement validity checks."""
    prs = new_presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title = slide.shapes.add_textbox(Inches(0.4), Inches(0.25), Inches(9.2), Inches(0.5))
    title.text_frame.text = "LibreOffice editor validity: text"
    title.text_frame.paragraphs[0].runs[0].font.size = Pt(18)

    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.5), Inches(7.8), Inches(1.8)
    )
    shape.name = "Editable Text Target"
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0xE2, 0xF0, 0xD9)
    shape.line.color.rgb = RGBColor(0x70, 0xAD, 0x47)
    shape.line.width = Pt(1.5)

    tf = shape.text_frame
    tf.word_wrap = True
    paragraph = tf.paragraphs[0]
    paragraph.text = text
    run = paragraph.runs[0]
    run.font.name = "Liberation Sans"
    run.font.size = Pt(30)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    prs.save(os.path.join(OUTPUT_DIR, filename))
    print(f"  Created: {filename}")


def create_editor_validity_transform_fixture(filename, left, top, width, height):
    """Fixture pair for editor-core move / resize validity checks."""
    prs = new_presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title = slide.shapes.add_textbox(Inches(0.4), Inches(0.25), Inches(9.2), Inches(0.5))
    title.text_frame.text = "LibreOffice editor validity: transform"
    title.text_frame.paragraphs[0].runs[0].font.size = Pt(18)

    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.name = "Move Resize Target"
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0x5B, 0x9B, 0xD5)
    shape.line.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    shape.line.width = Pt(2)

    tf = shape.text_frame
    paragraph = tf.paragraphs[0]
    paragraph.text = "Move + resize"
    run = paragraph.runs[0]
    run.font.name = "Liberation Sans"
    run.font.size = Pt(24)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    marker = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(7.7), Inches(3.9), Inches(0.45), Inches(0.45)
    )
    marker.fill.solid()
    marker.fill.fore_color.rgb = RGBColor(0xED, 0x7D, 0x31)
    marker.line.fill.background()

    prs.save(os.path.join(OUTPUT_DIR, filename))
    print(f"  Created: {filename}")


def create_editor_validity_formatting_fixture(filename, *, expected):
    """Fixture pair for editor-core run property validity checks."""
    prs = new_presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title = slide.shapes.add_textbox(Inches(0.4), Inches(0.25), Inches(9.2), Inches(0.5))
    title.text_frame.text = "LibreOffice editor validity: formatting"
    title.text_frame.paragraphs[0].runs[0].font.size = Pt(18)

    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(8.2), Inches(2.0)
    )
    shape.name = "Editable Formatting Target"
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0xFF, 0xF2, 0xCC)
    shape.line.color.rgb = RGBColor(0xBF, 0x90, 0x00)
    shape.line.width = Pt(1.5)

    tf = shape.text_frame
    tf.word_wrap = True
    paragraph = tf.paragraphs[0]
    paragraph.text = "Editable formatting target"
    paragraph.alignment = PP_ALIGN.CENTER
    run = paragraph.runs[0]

    if expected:
        run.font.name = "Liberation Serif"
        run.font.size = Pt(30)
        run.font.bold = False
        run.font.italic = False
        run.font.underline = False
        run.font.color.rgb = RGBColor(0x9C, 0x00, 0x00)
    else:
        run.font.name = "Liberation Sans"
        run.font.size = Pt(18)
        run.font.bold = True
        run.font.italic = True
        run.font.underline = True
        run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    prs.save(os.path.join(OUTPUT_DIR, filename))
    print(f"  Created: {filename}")


def create_editor_validity_paragraph_fixture(filename, *, expected):
    """Fixture pair for editor-core paragraph property validity checks."""
    prs = new_presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title = slide.shapes.add_textbox(Inches(0.4), Inches(0.25), Inches(9.2), Inches(0.5))
    title.text_frame.text = "LibreOffice editor validity: paragraph"
    title.text_frame.paragraphs[0].runs[0].font.size = Pt(18)

    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(1.3), Inches(8.0), Inches(2.0)
    )
    shape.name = "Editable Paragraph Target"
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0xDE, 0xEB, 0xF7)
    shape.line.color.rgb = RGBColor(0x2F, 0x75, 0xB5)
    shape.line.width = Pt(1.5)

    tf = shape.text_frame
    tf.word_wrap = True
    paragraph = tf.paragraphs[0]
    paragraph.text = "Paragraph properties target"
    paragraph.alignment = PP_ALIGN.RIGHT if expected else PP_ALIGN.LEFT
    paragraph.level = 1 if expected else 0
    if expected:
        p_pr = paragraph._p.get_or_add_pPr()
        bu_char = OxmlElement("a:buChar")
        bu_char.set("char", "\u2022")
        p_pr.insert(0, bu_char)
    run = paragraph.runs[0]
    run.font.name = "Liberation Sans"
    run.font.size = Pt(24)
    run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    prs.save(os.path.join(OUTPUT_DIR, filename))
    print(f"  Created: {filename}")


def create_editor_validity_image_fixture(filename, image_base64):
    """Fixture pair for editor-core image replacement validity checks."""
    prs = new_presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title = slide.shapes.add_textbox(Inches(0.4), Inches(0.25), Inches(9.2), Inches(0.5))
    title.text_frame.text = "LibreOffice editor validity: image"
    title.text_frame.paragraphs[0].runs[0].font.size = Pt(18)

    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        tmp.write(base64.b64decode(image_base64))
        image_path = tmp.name

    try:
        pic = slide.shapes.add_picture(
            image_path, Inches(2.2), Inches(1.4), Inches(4.8), Inches(2.7)
        )
        pic.name = "Replace Image Target"
    finally:
        os.unlink(image_path)

    prs.save(os.path.join(OUTPUT_DIR, filename))
    print(f"  Created: {filename}")


def create_editor_validity_fixtures():
    """PPTX source / expected pairs consumed by editor-validity.test.ts."""
    create_editor_validity_text_fixture(
        "editor-validity-text-source.pptx",
        "Original LibreOffice text",
    )
    create_editor_validity_text_fixture(
        "editor-validity-text-expected.pptx",
        "Edited LibreOffice text",
    )
    create_editor_validity_transform_fixture(
        "editor-validity-transform-source.pptx",
        Inches(0.9),
        Inches(1.2),
        Inches(2.4),
        Inches(1.2),
    )
    create_editor_validity_transform_fixture(
        "editor-validity-transform-expected.pptx",
        Inches(3.0),
        Inches(2.1),
        Inches(3.2),
        Inches(1.6),
    )
    create_editor_validity_formatting_fixture(
        "editor-validity-formatting-source.pptx",
        expected=False,
    )
    create_editor_validity_formatting_fixture(
        "editor-validity-formatting-expected.pptx",
        expected=True,
    )
    create_editor_validity_paragraph_fixture(
        "editor-validity-paragraph-source.pptx",
        expected=False,
    )
    create_editor_validity_paragraph_fixture(
        "editor-validity-paragraph-expected.pptx",
        expected=True,
    )
    create_editor_validity_image_fixture(
        "editor-validity-image-source.pptx",
        "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAEUlEQVR4nGP8z4AATEhsPBwAM9EBBzDn4UwAAAAASUVORK5CYII=",
    )
    create_editor_validity_image_fixture(
        "editor-validity-image-expected.pptx",
        "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAE0lEQVR4nGNkYPjPAANMcBZeDgAx0wEH1s7nlgAAAABJRU5ErkJggg==",
    )


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print("Generating editor-validity fixtures...")
    create_basic_shapes()
    create_editor_validity_fixtures()
    print("Done!")


if __name__ == "__main__":
    main()

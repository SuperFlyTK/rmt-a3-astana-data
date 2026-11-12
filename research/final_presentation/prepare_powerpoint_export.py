"""Make the two Canva result tables native PowerPoint tables for submission.

Pass the PPTX downloaded from Canva and an output path. The PowerPoint copy
uses common fonts so that readers without Canva's fonts keep readable layouts.
"""

import argparse
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Emu, Pt


SOURCE = "https://github.com/SuperFlyTK/rmt-a3-astana-data/blob/4c956a3/research/assignment4/"
NAVY = RGBColor.from_string("18265A")
WHITE = RGBColor.from_string("FFFFFF")
PAPER = RGBColor.from_string("F7F7F5")
ROWS_EXPECTATIONS = [
    ("Check", "Falsified if", "ORS check results", "Verdict"),
    ("Route validity", "≥2% fail", "0/572 failed (0%)", "Supported"),
    ("API response", "<95% within 10 s", "572/572 within 10 s", "Supported"),
    ("Pilot coverage", "<20 valid pilot routes", "20/20 in all 3 cycles", "Supported"),
    ("Scheduled coverage", "<692 automatic attempts", "408/768 (53.125%)", "Missed 90% target"),
]
ROWS_OBJECTIVES = [
    ("Objective", "Verdict", "Finding"),
    ("Routes", "Partial", "20 routes, fixed periods. Representativeness untested."),
    ("Collection", "Missed", "80 Google–ORS pairs. 14-day Google–Yandex plan unmet."),
    ("Metrics", "Partial", "Signed-gap IQRs 218.30–279.48 s. No confidence intervals."),
    ("Periods", "Partial", "271.18 s descriptive gap. Route-level inference unavailable."),
    ("Patterns", "Partial", "Window medians 379.40–748.85 s. Recurring route patterns unresolved."),
]


def walk(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from walk(shape.shapes)


def apply_portable_fonts(presentation):
    for index, slide in enumerate(presentation.slides):
        for shape in walk(slide.shapes):
            if not shape.has_text_frame or not shape.text.strip():
                continue
            compact = index in (2, 3) and len(shape.text.split()) > 14
            reference_heading = index >= 11 and shape.text.startswith("Earlier question ·")
            for paragraph in shape.text_frame.paragraphs:
                if compact:
                    paragraph.line_spacing = 1.1 if paragraph.text.strip() else Pt(19)
                elif index == 10 and shape.text.startswith(("10 weeks of shared work", "Ten-week team workflow")):
                    paragraph.line_spacing = 1.15
                for run in paragraph.runs:
                    if reference_heading and run.font.size and run.font.size > Pt(72):
                        run.font.size = Pt(72)
                    name = run.font.name or ""
                    if name.startswith("Roca Two"):
                        run.font.name = "Georgia"
                        run.font.bold = bool(run.font.bold or "Bold" in name)
                    elif name.startswith("TT Chocolates"):
                        run.font.name = "Arial"
                        run.font.bold = bool(run.font.bold or "Bold" in name)


def add_cell_borders(cell):
    properties = cell._tc.get_or_add_tcPr()
    for side in ("lnL", "lnR", "lnT", "lnB"):
        line = OxmlElement(f"a:{side}")
        line.set("w", str(Pt(0.7)))
        fill = OxmlElement("a:solidFill")
        color = OxmlElement("a:srgbClr")
        color.set("val", "CFD4DC")
        fill.append(color)
        line.append(fill)
        properties.append(line)


def replace_exported_table(slide, rows, widths, top, height, *, source_link=False):
    exported = [shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT]
    if len(exported) != 1:
        raise ValueError("Expected one Canva table exported as an embedded object on this slide.")
    element = exported[0].element
    element.getparent().remove(element)
    pixels = 9525  # Canva's 1920 × 1080 design in the exported PPTX.
    table_shape = slide.shapes.add_table(
        len(rows), len(widths), Emu(135 * pixels), Emu(top * pixels),
        Emu(sum(widths) * pixels), Emu(height * pixels),
    )
    if source_link:
        table_shape.click_action.hyperlink.address = SOURCE + "tables/expectations.csv"
    table = table_shape.table
    table.first_row = False
    table.horz_banding = False
    for column, width in zip(table.columns, widths):
        column.width = Emu(width * pixels)
    for row, values in enumerate(rows):
        for column, value in enumerate(values):
            cell = table.cell(row, column)
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY if row == 0 else PAPER
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Pt(10)
            cell.margin_top = cell.margin_bottom = Pt(5)
            cell.text_frame.word_wrap = True
            para = cell.text_frame.paragraphs[0]
            para.line_spacing = 1.12
            para.space_before = para.space_after = Pt(0)
            run = para.add_run()
            run.text = value
            run.font.name = "Arial"
            run.font.size = Pt(29)
            run.font.bold = row == 0
            run.font.color.rgb = WHITE if row == 0 else NAVY
            add_cell_borders(cell)


def link_results(slide):
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text.strip() == "+271.18 s":
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if run.hyperlink.address:
                        run.hyperlink.address = None
            shape.click_action.hyperlink.address = SOURCE + "figures/02_period_comparison.png"
        elif abs(shape.left - 780 * 9525) < 10000 and abs(shape.width - 1040 * 9525) < 10000:
            shape.click_action.hyperlink.address = SOURCE + "figures/01_window_disagreement.png"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("canva_export", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    prs = Presentation(args.canva_export)
    if len(prs.slides) < 11 or abs(prs.slide_width - 1920 * 9525) > 100:
        raise ValueError("Expected the main 11-slide Canva defence in 1920 × 1080 format.")
    apply_portable_fonts(prs)
    replace_exported_table(prs.slides[7], ROWS_EXPECTATIONS, [415] * 4, 329, 620, source_link=True)
    replace_exported_table(prs.slides[8], ROWS_OBJECTIVES, [340, 250, 1070], 314, 677)
    link_results(prs.slides[5])
    prs.save(args.output)
    print(args.output)


if __name__ == "__main__":
    main()

"""Make a plain white Times New Roman copy of the 16-slide defence export."""

import argparse
from hashlib import sha256
from io import BytesIO
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Emu

from plot_study_map import draw_map
from plot_window_results import draw_chart
from prepare_powerpoint_export import add_cell_borders


FONT = "Times New Roman"
BLACK = RGBColor.from_string("000000")
WHITE = RGBColor.from_string("FFFFFF")
FIGURES = {
    7: ("58a6de3d9e639ea128dfb33479f038ca1cf86d88c2adedb4ab60470a4fb354a9", draw_map),
    8: ("6c7063a8d273922c8b842287036e24fcb1b4dee03c042e389ddfe4ceb5d75b65", draw_chart),
}


def style_text(frame, *, visible=True):
    for paragraph in frame.paragraphs:
        paragraph.font.name = FONT
        paragraph.font._element.attrib.pop("spc", None)
        for run in paragraph.runs:
            run.font.name = FONT
            run.font._element.attrib.pop("spc", None)
            if visible:
                run.font.color.rgb = BLACK


def style_table(shape):
    table = shape.table
    table.first_row = table.horz_banding = False
    # Reserve room for the source line beneath the table.
    row_height = Emu(570 * 9525 // len(table.rows))
    for row_index, row in enumerate(table.rows):
        row.height = row_height
        for cell in row.cells:
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE
            style_text(cell.text_frame)
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = row_index == 0
            for border in cell._tc.xpath("./a:tcPr/*[starts-with(local-name(), 'ln')]"):
                border.getparent().remove(border)
            add_cell_borders(cell)


def update_figure(slide, shape, expected_hash, draw):
    blips = shape.element.xpath(".//a:blip[@r:embed]")
    if not blips:
        return False
    part = slide.part.related_part(blips[0].get(qn("r:embed")))
    if sha256(part.blob).hexdigest() != expected_hash:
        return False
    image = draw(font="times.ttf", bold_font="timesbd.ttf", paper="white")
    with BytesIO() as buffer:
        image.save(buffer, format="PNG", dpi=(144, 144))
        part.blob = buffer.getvalue()
    return True


def simplify_slide(slide, number):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = WHITE
    slide.element.set("showMasterSp", "0")
    figure = FIGURES.get(number)
    kept_figure = False
    orphan_candidates = set()
    relationship_ns = qn("r:id").split("}")[0] + "}"
    for shape in list(slide.shapes):
        if shape.has_text_frame and shape.text.strip():
            shape.fill.background()
            shape.line.fill.background()
            style_text(shape.text_frame)
        elif shape.has_table:
            style_table(shape)
        elif figure and update_figure(slide, shape, *figure):
            kept_figure = True
        else:
            for element in shape.element.iter():
                orphan_candidates.update(
                    value for key, value in element.attrib.items() if key.startswith(relationship_ns)
                )
            shape.element.getparent().remove(shape.element)
    if figure and not kept_figure:
        raise ValueError(f"Slide {number} is missing its expected research figure.")
    live_references = {
        value
        for element in slide.element.iter()
        for key, value in element.attrib.items()
        if key.startswith(relationship_ns)
    }
    for rid in (orphan_candidates - live_references) & slide.part.rels.keys():
        slide.part.rels.pop(rid)
    if slide.has_notes_slide:
        frame = slide.notes_slide.notes_text_frame
        if frame is not None:
            style_text(frame, visible=False)


def update_theme_fonts(presentation):
    for part in presentation.part.package.iter_parts():
        if not str(part.partname).startswith("/ppt/theme/"):
            continue
        theme = etree.fromstring(part.blob)
        for font in theme.xpath(".//*[@typeface]"):
            font.set("typeface", FONT)
        part.blob = etree.tostring(theme, encoding="UTF-8", xml_declaration=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    if args.input.resolve() == args.output.resolve():
        parser.error("Choose a different output path to preserve the original.")
    presentation = Presentation(args.input)
    if len(presentation.slides) != 16 or presentation.slide_width != 1920 * 9525:
        parser.error("Expected the 16-slide Astana defence export.")
    for number, slide in enumerate(presentation.slides, 1):
        simplify_slide(slide, number)
    update_theme_fonts(presentation)
    presentation.save(args.output)
    print(args.output)


if __name__ == "__main__":
    main()

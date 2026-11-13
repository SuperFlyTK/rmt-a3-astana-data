"""Render the archived window medians with labels sized for projection."""

import csv
from datetime import datetime, timedelta, timezone
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


HERE = Path(__file__).resolve().parent
SOURCE = HERE.parent / "assignment4" / "tables" / "batch_results.csv"
OUTPUT = HERE / "output" / "Window_results_projector.png"
INK = "#18265a"
PAPER = "#fbfbf9"


def draw_chart(*, font="arial.ttf", bold_font="arialbd.ttf", paper=PAPER):
    with SOURCE.open(encoding="utf-8", newline="") as stream:
        windows = [row for row in csv.DictReader(stream) if row["phase"] == "followup"]
    assert len(windows) == 4
    assert sum(int(row["n_valid_pairs"]) for row in windows) == 80
    assert [row["period"] for row in windows] == ["peak", "offpeak", "offpeak", "peak"]

    canvas = Image.new("RGB", (1060, 610), paper)
    draw = ImageDraw.Draw(canvas)
    fonts = Path("C:/Windows/Fonts")
    label = ImageFont.truetype(str(fonts / bold_font), 39)
    value = ImageFont.truetype(str(fonts / bold_font), 43)
    caption = ImageFont.truetype(str(fonts / font), 35)
    title = ImageFont.truetype(str(fonts / bold_font), 43)
    draw.text((32, 24), "Window-by-window results", font=title, fill=INK)
    draw.text((32, 78), "Median absolute gap (minutes)", font=caption, fill=INK)

    baseline = 467
    scale = 22
    draw.line((35, baseline, 1025, baseline), fill="#7a859e", width=3)
    local = timezone(timedelta(hours=5))
    for index, row in enumerate(windows):
        center = 155 + 250 * index
        minutes = float(row["median_absolute_difference_s"]) / 60
        bar_top = baseline - minutes * scale
        peak = row["period"] == "peak"
        date = datetime.fromisoformat(row["start_utc"]).astimezone(local)
        draw.rectangle(
            (center - 71, bar_top, center + 71, baseline),
            fill=INK if peak else "#607b9a",
        )
        draw.text((center, bar_top - 12), f"{minutes:.2f}", font=value, fill=INK, anchor="mb")
        draw.text((center, 501), f"{date.day} Oct", font=label, fill=INK, anchor="mt")
        draw.text((center, 557), "Peak" if peak else "Off-peak", font=caption, fill=INK, anchor="mt")

    return canvas


def main():
    OUTPUT.parent.mkdir(exist_ok=True)
    draw_chart().save(OUTPUT, dpi=(144, 144))
    print(OUTPUT)


if __name__ == "__main__":
    main()

"""Draw the verified study coordinates and archived ORS pilot paths for Canva."""

import csv
import json
import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


HERE = Path(__file__).resolve().parent
DATA = HERE.parent / "assignment4" / "data"
OUTPUT = HERE / "output" / "Astana_study_map.png"
LABELS = {
    "P01": ("Khan Shatyr", "left", 245),
    "P02": ("Airport T1", "right", 651),
    "P03": ("EXPO centre", "right", 428),
    "P04": ("Saryarka Ave", "left", 132),
    "P05": ("Abay Ave", "right", 72),
    "P06": ("Bogenbay Batyr Ave", "left", 57),
    "P07": ("Kenesary St", "right", 155),
    "P08": ("Syganak St", "left", 329),
    "P09": ("Nurly Zhol station", "right", 330),
    "P10": ("EXPO Museum", "left", 514),
}


def rows(name):
    with (DATA / name).open(encoding="utf-8", newline="") as stream:
        return list(csv.DictReader(stream))


def local_km(lon, lat):
    return (
        (lon - 71.45) * 111.195 * math.cos(math.radians(51.1)),
        (lat - 51.1) * 111.195,
    )


def main():
    endpoints = {
        row["endpoint_id"]: (float(row["lon"]), float(row["lat"]))
        for row in rows("selected_endpoints.csv")
    }
    routes = {
        row["route_id"]: (
            (float(row["origin_lon"]), float(row["origin_lat"])),
            (float(row["destination_lon"]), float(row["destination_lat"])),
        )
        for row in rows("routes.csv")
    }
    archive = json.loads((DATA / "ors_pilot_routes.geojson").read_text(encoding="utf-8"))
    paths = {
        feature["properties"]["route_id"]: feature["geometry"]["coordinates"]
        for feature in archive["features"]
    }
    pairs = {frozenset(pair) for pair in routes.values()}
    used = {point for pair in routes.values() for point in pair}
    assert len(endpoints) == 10 and set(endpoints) == set(LABELS)
    assert used == set(endpoints.values())
    assert len(routes) == 20 and len(pairs) == 10 and routes.keys() == paths.keys()
    assert len(set(routes.values())) == 20
    assert all((end, start) in routes.values() for start, end in routes.values())
    snap_distances = [
        1000 * math.dist(local_km(*given), local_km(*snapped))
        for route, (start, end) in routes.items()
        for given, snapped in [(start, paths[route][0]), (end, paths[route][-1])]
    ]

    canvas = Image.new("RGB", (1780, 760), "#fbfbf9")
    draw = ImageDraw.Draw(canvas)
    ink, route_color, accent = "#18265a", "#7b8ba6", "#ac3d19"
    fonts = Path("C:/Windows/Fonts")
    label_font = ImageFont.truetype(str(fonts / "arialbd.ttf"), 37)
    body_font = ImageFont.truetype(str(fonts / "arial.ttf"), 38)
    small_font = ImageFont.truetype(str(fonts / "arial.ttf"), 29)
    count_font = ImageFont.truetype(str(fonts / "arialbd.ttf"), 59)
    scale = 36

    def to_canvas(point):
        east, north = local_km(*point)
        return 674 + scale * east, 370 - scale * north

    for path in paths.values():
        draw.line([to_canvas(point) for point in path], fill=route_color, width=5)
    draw.line([to_canvas(point) for point in paths["R01"]], fill=accent, width=11)
    for endpoint_id, (label, side, y_label) in LABELS.items():
        x, y = to_canvas(endpoints[endpoint_id])
        end_x = 412 if side == "left" else 927
        start_x = end_x - 13 if side == "left" else end_x + 13
        anchor = "rm" if side == "left" else "lm"
        draw.line([(x, y), (end_x, y_label)], fill="#525d74", width=3)
        draw.text((start_x, y_label), label, font=label_font, fill=ink, anchor=anchor)
    for point in endpoints.values():
        x, y = to_canvas(point)
        draw.ellipse((x - 13, y - 13, x + 13, y + 13), fill=ink, outline="white", width=4)

    draw.line((43, 643, 43, 588), fill=ink, width=4)
    draw.polygon([(43, 581), (33, 601), (53, 601)], fill=ink)
    draw.text((43, 568), "N", font=small_font, fill=ink, anchor="mb")
    draw.line((93, 642, 93 + 2 * scale, 642), fill=ink, width=5)
    draw.text((93, 649), "2 km", font=small_font, fill=ink, anchor="lt")
    draw.line((1301, 67, 1301, 665), fill="#ccd3df", width=3)
    draw.text((1340, 71), "10 points", font=count_font, fill=ink)
    draw.text((1341, 145), "20 directed routes", font=body_font, fill=ink)
    draw.line((1342, 264, 1400, 264), fill=accent, width=11)
    draw.text((1341, 286), "Example pilot path", font=label_font, fill=ink)
    draw.text((1341, 340), "Khan Shatyr\nto Airport T1", font=body_font, fill=ink, spacing=8)
    draw.text((1341, 492), "Public roadside points\n24–25 Oct 2026", font=body_font, fill=ink, spacing=8)
    draw.text((40, 721), "Endpoint data © OpenStreetMap contributors. Paths from the openrouteservice pilot.",
              font=small_font, fill=ink, anchor="lm")
    OUTPUT.parent.mkdir(exist_ok=True)
    canvas.save(OUTPUT, dpi=(144, 144))
    print(f"{len(endpoints)} endpoints, {len(pairs)} reciprocal pairs, {len(paths)} archived paths.")
    print(f"Maximum pilot endpoint snap: {max(snap_distances):.1f} m.")
    print(OUTPUT)


if __name__ == "__main__":
    main()

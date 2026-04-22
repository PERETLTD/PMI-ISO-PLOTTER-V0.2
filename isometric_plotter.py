#!/usr/bin/env python3
from __future__ import annotations

import argparse
import io
import json
import re
from pathlib import Path
from typing import Iterable

from pypdf import PdfReader, PdfWriter
from pypdf.generic import DecodedStreamObject, NameObject
from reportlab.lib import colors
from reportlab.lib.colors import Color
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


PAGE_WIDTH, PAGE_HEIGHT = A4

# Measured from the supplied PDF template.
GRID_ORIGIN_TOP_LEFT = (41.038108, 36.0)
GRID_AXIS_U = (0.0, 28.3464567)  # one cell straight down on a vertical grid line
GRID_AXIS_V = (24.549, 14.17322835)  # one cell down-right on a 30 degree grid line

DEFAULT_TEMPLATE = Path("/Users/thisaruperera/Downloads/print-graph-paper.com.pdf")
DEFAULT_OUTPUT = Path("branded_isometric_plot.pdf")


def hex_to_color(value: str, default: str = "#1f4b99") -> Color:
    try:
        return colors.HexColor(value)
    except Exception:
        return colors.HexColor(default)


def grid_to_page(point: Iterable[float]) -> tuple[float, float]:
    u, v = point
    x = GRID_ORIGIN_TOP_LEFT[0] + (u * GRID_AXIS_U[0]) + (v * GRID_AXIS_V[0])
    y_top = GRID_ORIGIN_TOP_LEFT[1] + (u * GRID_AXIS_U[1]) + (v * GRID_AXIS_V[1])
    return x, PAGE_HEIGHT - y_top


def remove_source_footer(page) -> None:
    content_object = page.get_contents()
    content = content_object.get_data()

    exact_block = (
        b"q\n1 0 0 1 0 758.4 cm\nq\nBT\n0.733 0.733 0.733 rg\n/GS1 gs\n/F0 -10 Tf\n"
        b"0 9.8809 Td (print-graph-paper.com) Tj\nET\nQ"
    )
    if exact_block in content:
        cleaned = content.replace(exact_block, b"")
    else:
        cleaned = re.sub(rb".{0,120}\(print-graph-paper\.com\) Tj.{0,80}", b"", content, count=1, flags=re.S)

    stream = DecodedStreamObject()
    stream.set_data(cleaned)
    try:
        stream[NameObject("/Filter")] = content_object.get(NameObject("/Filter"))  # type: ignore[index]
    except Exception:
        pass
    page.replace_contents(stream)


def draw_polyline(pdf: canvas.Canvas, points: list[list[float]], stroke: str, width: float) -> None:
    if len(points) < 2:
        return
    path = pdf.beginPath()
    x0, y0 = grid_to_page(points[0])
    path.moveTo(x0, y0)
    for point in points[1:]:
        x, y = grid_to_page(point)
        path.lineTo(x, y)
    pdf.setStrokeColor(hex_to_color(stroke))
    pdf.setLineWidth(width)
    pdf.setLineCap(1)
    pdf.setLineJoin(1)
    pdf.drawPath(path, stroke=1, fill=0)


def draw_segments(pdf: canvas.Canvas, segments: list[dict]) -> None:
    for segment in segments:
        start = segment["from"]
        end = segment["to"]
        stroke = segment.get("stroke", "#d9480f")
        width = float(segment.get("width", 2))
        draw_polyline(pdf, [start, end], stroke, width)


def draw_polylines(pdf: canvas.Canvas, polylines: list[dict]) -> None:
    for polyline in polylines:
        draw_polyline(
            pdf,
            polyline["points"],
            polyline.get("stroke", "#d9480f"),
            float(polyline.get("width", 2)),
        )


def draw_logo(pdf: canvas.Canvas, x: float = 18, y: float = 14, width: float = 112) -> None:
    height = width * 0.61
    base = pdf.beginPath()
    base.moveTo(x + width * 0.10, y)
    base.lineTo(x + width * 0.90, y)
    base.lineTo(x + width, y + height * 0.12)
    base.lineTo(x + width, y + height * 0.90)
    base.lineTo(x + width * 0.92, y + height)
    base.lineTo(x + width * 0.10, y + height)
    base.lineTo(x, y + height * 0.88)
    base.lineTo(x, y + height * 0.12)
    base.close()

    pdf.setFillColor(colors.HexColor("#1F3F96"))
    pdf.setStrokeColor(colors.HexColor("#1F3F96"))
    pdf.drawPath(base, stroke=0, fill=1)

    pdf.setFillColor(colors.white)
    pdf.setFont("Helvetica-Bold", height * 0.56)
    pdf.drawString(x + width * 0.12, y + height * 0.49, "Pmi")

    pdf.setFont("Helvetica", height * 0.18)
    pdf.drawCentredString(x + width * 0.52, y + height * 0.30, "Premier")
    pdf.setFont("Helvetica-Bold", height * 0.16)
    pdf.drawCentredString(x + width * 0.52, y + height * 0.14, "Mechanical Installations")


def create_overlay(spec: dict) -> bytes:
    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)

    for segment in spec.get("segments", []):
        draw_segments(pdf, [segment])

    draw_polylines(pdf, spec.get("polylines", []))

    branding = spec.get("branding", {})
    if branding.get("enabled", True):
        draw_logo(
            pdf,
            x=float(branding.get("x", 18)),
            y=float(branding.get("y", 14)),
            width=float(branding.get("width", 112)),
        )

    pdf.showPage()
    pdf.save()
    return buffer.getvalue()


def load_spec(path: Path | None) -> dict:
    if path is None:
        return {}
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def build_pdf(template_path: Path, spec_path: Path | None, output_path: Path) -> None:
    writer = PdfWriter(clone_from=str(template_path))
    source_page = writer.pages[0]
    remove_source_footer(source_page)

    overlay_reader = PdfReader(io.BytesIO(create_overlay(load_spec(spec_path))))
    source_page.merge_page(overlay_reader.pages[0])

    with output_path.open("wb") as handle:
        writer.write(handle)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Brand and plot onto the supplied isometric graph paper template."
    )
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE, help="Path to the source PDF template.")
    parser.add_argument("--spec", type=Path, help="JSON file describing plot segments and polylines.")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Path for the generated PDF.")
    args = parser.parse_args()

    build_pdf(args.template, args.spec, args.output)


if __name__ == "__main__":
    main()

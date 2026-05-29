#!/usr/bin/env python3
"""Genera imágenes de laboratorio (fondo ADN + esquina) para la plantilla qPCR."""
from __future__ import annotations

import math
from pathlib import Path

from PIL import Image, ImageDraw

ASSETS = Path(__file__).resolve().parent / "assets"
W, H = 960, 72
BG = (224, 242, 241)
TEAL = (0, 121, 107)
TEAL_LIGHT = (128, 203, 196)
WHITE = (255, 255, 255)


def helix(draw: ImageDraw.ImageDraw, y_mid: float, amp: float, phase: float, x0: int, x1: int) -> None:
    step = 14
    pts_a: list[tuple[int, int]] = []
    pts_b: list[tuple[int, int]] = []
    for x in range(x0, x1, step):
        t = (x - x0) / max(x1 - x0, 1) * math.pi * 5 + phase
        ya = int(y_mid + amp * math.sin(t))
        yb = int(y_mid + amp * math.sin(t + math.pi))
        pts_a.append((x, ya))
        pts_b.append((x, yb))
        draw.ellipse((x - 3, ya - 3, x + 3, ya + 3), fill=TEAL_LIGHT + (80,))
        draw.ellipse((x - 3, yb - 3, x + 3, yb + 3), fill=TEAL + (100,))
    if len(pts_a) > 1:
        draw.line(pts_a, fill=TEAL, width=2)
        draw.line(pts_b, fill=TEAL_LIGHT, width=2)
    for i in range(0, len(pts_a) - 1, 2):
        draw.line([pts_a[i], pts_b[i]], fill=TEAL_LIGHT, width=1)


def main() -> None:
    ASSETS.mkdir(parents=True, exist_ok=True)
    img = Image.new("RGBA", (W, H), BG + (255,))
    draw = ImageDraw.Draw(img)
    helix(draw, H * 0.38, 10, 0.0, 0, W)
    helix(draw, H * 0.62, 8, 1.2, 40, W - 20)
    for x in range(20, W, 48):
        draw.rectangle((x, 4, x + 2, H - 4), fill=WHITE + (40,))
    out = ASSETS / "fondo_cabecera_raw.png"
    img.save(out, "PNG")
    print(f"Wrote {out} ({out.stat().st_size} bytes)")

    corner = Image.new("RGBA", (120, 120), (0, 0, 0, 0))
    cd = ImageDraw.Draw(corner)
    helix(cd, 60, 18, 0.5, 0, 120)
    corner.save(ASSETS / "dna_esquina.png", "PNG")
    print(f"Wrote {ASSETS / 'dna_esquina.png'}")


if __name__ == "__main__":
    main()

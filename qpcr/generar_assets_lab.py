#!/usr/bin/env python3
"""Genera imágenes de laboratorio / biología molecular para la plantilla qPCR.

Produce, con anti-aliasing (supersampling + LANCZOS):
  - fondo_cabecera_raw.png : banner doble hélice nítido para la cabecera de RAW
  - dna_esquina.png        : motivo molecular transparente para esquinas de hojas

Uso:  python3 qpcr/generar_assets_lab.py
"""
from __future__ import annotations

import math
from pathlib import Path

from PIL import Image, ImageDraw

ASSETS = Path(__file__).resolve().parent / "assets"

# Paleta laboratorio (coherente con la plantilla / macro)
TEAL = (0, 105, 92)          # #00695C acento
TEAL_DARK = (0, 77, 64)      # #004D40
TEAL_MID = (0, 137, 123)     # #00897B
TEAL_LIGHT = (77, 182, 172)  # #4DB6AC
MINT = (128, 203, 196)       # #80CBC4
HEADER_TOP = (234, 247, 245)  # degradado claro
HEADER_BOT = (210, 236, 232)  # degradado algo más intenso
WHITE = (255, 255, 255)

SS = 4  # factor de supersampling


def _lerp(a: tuple[int, int, int], b: tuple[int, int, int], t: float) -> tuple[int, int, int]:
    t = max(0.0, min(1.0, t))
    return tuple(int(round(a[i] + (b[i] - a[i]) * t)) for i in range(3))


def _vertical_gradient(w: int, h: int, top: tuple[int, int, int], bot: tuple[int, int, int]) -> Image.Image:
    img = Image.new("RGBA", (w, h))
    px = img.load()
    for y in range(h):
        c = _lerp(top, bot, y / max(h - 1, 1)) + (255,)
        for x in range(w):
            px[x, y] = c
    return img


def _draw_strand(
    draw: ImageDraw.ImageDraw,
    x0: int,
    x1: int,
    y_mid: float,
    amp: float,
    cycles: float,
    phase: float,
    width: int,
    c_a: tuple[int, int, int],
    c_b: tuple[int, int, int],
) -> list[tuple[int, int]]:
    """Dibuja una hebra senoidal con color interpolado; devuelve los puntos."""
    pts: list[tuple[int, int]] = []
    n = max(x1 - x0, 1)
    samples = max(int(n / 2), 24)
    for i in range(samples + 1):
        f = i / samples
        x = int(x0 + f * n)
        t = f * math.pi * 2 * cycles + phase
        y = int(round(y_mid + amp * math.sin(t)))
        pts.append((x, y))
    for i in range(len(pts) - 1):
        col = _lerp(c_a, c_b, i / max(len(pts) - 1, 1))
        draw.line([pts[i], pts[i + 1]], fill=col + (255,), width=width)
    return pts


def _double_helix(
    draw: ImageDraw.ImageDraw,
    x0: int,
    x1: int,
    y_mid: float,
    amp: float,
    cycles: float,
    strand_w: int,
    rung_w: int,
    node_r: int,
    rung_step: int = 18,
) -> None:
    """Dos hebras desfasadas π con peldaños (pares de bases) y nodos."""
    n = max(x1 - x0, 1)

    def y_of(f: float, phase: float) -> float:
        return y_mid + amp * math.sin(f * math.pi * 2 * cycles + phase)

    # Peldaños primero (quedan por detrás de las hebras)
    f = 0.0
    k = 0
    while True:
        x = int(x0 + f * n)
        if x > x1:
            break
        ya = y_of(f, 0.0)
        yb = y_of(f, math.pi)
        # color del peldaño según fase (efecto par de bases)
        col = TEAL_MID if (k % 2 == 0) else TEAL_LIGHT
        draw.line([(x, int(ya)), (x, int(yb))], fill=col + (235,), width=rung_w)
        f += rung_step / n
        k += 1

    a = _draw_strand(draw, x0, x1, y_mid, amp, cycles, 0.0, strand_w, TEAL_DARK, TEAL_MID)
    b = _draw_strand(draw, x0, x1, y_mid, amp, cycles, math.pi, strand_w, TEAL_LIGHT, MINT)

    # Nodos (nucleótidos) sobre las hebras
    for pts, fill in ((a, TEAL), (b, MINT)):
        step = max(len(pts) // int(cycles * 8 + 1), 3)
        for i in range(0, len(pts), step):
            x, y = pts[i]
            draw.ellipse(
                (x - node_r, y - node_r, x + node_r, y + node_r),
                fill=fill + (255,),
                outline=WHITE + (220,),
                width=max(node_r // 3, 1),
            )


def make_banner() -> None:
    w, h = 1400, 84
    img = _vertical_gradient(w * SS, h * SS, HEADER_TOP, HEADER_BOT)
    draw = ImageDraw.Draw(img, "RGBA")

    pad = 12 * SS
    _double_helix(
        draw,
        x0=pad,
        x1=w * SS - pad,
        y_mid=h * SS * 0.5,
        amp=h * SS * 0.30,
        cycles=9.0,
        strand_w=3 * SS,
        rung_w=2 * SS,
        node_r=3 * SS,
        rung_step=14 * SS,
    )

    # Línea base teal sutil (asienta el banner)
    draw.line([(0, h * SS - SS), (w * SS, h * SS - SS)], fill=TEAL + (180,), width=2 * SS)

    img = img.resize((w, h), Image.LANCZOS)
    out = ASSETS / "fondo_cabecera_raw.png"
    img.save(out, "PNG")
    print(f"Wrote {out} ({out.stat().st_size} bytes, {img.size[0]}x{img.size[1]})")


def make_corner() -> None:
    """Motivo molecular vertical, transparente, para esquina de cabeceras."""
    w, h = 132, 132
    img = Image.new("RGBA", (w * SS, h * SS), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img, "RGBA")

    # Hélice vertical: intercambiamos ejes dibujando en horizontal y rotando
    tmp = Image.new("RGBA", (h * SS, w * SS), (0, 0, 0, 0))
    td = ImageDraw.Draw(tmp, "RGBA")
    _double_helix(
        td,
        x0=6 * SS,
        x1=h * SS - 6 * SS,
        y_mid=w * SS * 0.5,
        amp=w * SS * 0.26,
        cycles=2.2,
        strand_w=3 * SS,
        rung_w=2 * SS,
        node_r=4 * SS,
        rung_step=16 * SS,
    )
    img = tmp.rotate(90, expand=True)

    img = img.resize((w, h), Image.LANCZOS)
    out = ASSETS / "dna_esquina.png"
    img.save(out, "PNG")
    print(f"Wrote {out} ({out.stat().st_size} bytes, {img.size[0]}x{img.size[1]})")


def main() -> None:
    ASSETS.mkdir(parents=True, exist_ok=True)
    make_banner()
    make_corner()


if __name__ == "__main__":
    main()

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter


SIZE = 1024
BG_TOP = (12, 24, 36, 255)
BG_BOTTOM = (28, 57, 82, 255)
GOLD = (199, 151, 52, 255)
GOLD_GLOW = (236, 196, 108, 160)
PAPER = (245, 238, 225, 255)
PAPER_EDGE = (219, 201, 172, 255)
INK = (20, 36, 52, 255)
INK_SOFT = (52, 74, 95, 255)


def lerp(start: tuple[int, int, int, int], end: tuple[int, int, int, int], factor: float) -> tuple[int, int, int, int]:
    return tuple(int(start[index] + (end[index] - start[index]) * factor) for index in range(4))


def build_gradient(size: int) -> Image.Image:
    image = Image.new("RGBA", (size, size))
    pixels = image.load()
    for y in range(size):
        factor = y / max(size - 1, 1)
        color = lerp(BG_TOP, BG_BOTTOM, factor)
        for x in range(size):
            pixels[x, y] = color
    return image


def add_glow(base: Image.Image) -> Image.Image:
    glow = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(glow)
    size = base.size[0]
    draw.ellipse((size * 0.38, -size * 0.04, size * 1.02, size * 0.56), fill=GOLD_GLOW)
    draw.ellipse((-size * 0.10, size * 0.52, size * 0.46, size * 1.08), fill=(87, 126, 156, 90))
    glow = glow.filter(ImageFilter.GaussianBlur(size // 11))
    return Image.alpha_composite(base, glow)


def draw_frame(base: Image.Image) -> None:
    draw = ImageDraw.Draw(base)
    inset = 58
    draw.rounded_rectangle(
        (inset, inset, SIZE - inset, SIZE - inset),
        radius=156,
        outline=(241, 214, 151, 180),
        width=8,
    )
    draw.rounded_rectangle(
        (inset + 24, inset + 24, SIZE - inset - 24, SIZE - inset - 24),
        radius=138,
        outline=(68, 94, 119, 210),
        width=4,
    )


def make_sheet() -> Image.Image:
    sheet = Image.new("RGBA", (430, 540), (0, 0, 0, 0))
    draw = ImageDraw.Draw(sheet)
    draw.rounded_rectangle((0, 0, 430, 540), radius=54, fill=PAPER, outline=PAPER_EDGE, width=8)
    draw.rounded_rectangle((52, 56, 378, 112), radius=18, fill=GOLD)
    for y in (168, 226, 284, 342):
        draw.rounded_rectangle((58, y, 312, y + 16), radius=8, fill=INK_SOFT)
    draw.rounded_rectangle((58, 402, 254, 434), radius=12, fill=(214, 182, 112, 255))
    draw.rounded_rectangle((278, 402, 350, 434), radius=12, fill=(219, 201, 172, 255))
    draw.ellipse((298, 180, 360, 242), outline=GOLD, width=10)
    return sheet


def add_documents(base: Image.Image) -> None:
    shadow = Image.new("RGBA", base.size, (0, 0, 0, 0))
    sheet = make_sheet()

    back_sheet = sheet.rotate(-8, expand=True)
    front_sheet = sheet.rotate(6, expand=True)

    shadow.paste((0, 0, 0, 110), (264, 248), back_sheet)
    shadow.paste((0, 0, 0, 140), (338, 188), front_sheet)
    shadow = shadow.filter(ImageFilter.GaussianBlur(22))
    base.alpha_composite(shadow)

    base.alpha_composite(back_sheet, (252, 232))
    base.alpha_composite(front_sheet, (326, 170))


def add_magnifier(base: Image.Image) -> None:
    lens = Image.new("RGBA", (320, 320), (0, 0, 0, 0))
    draw = ImageDraw.Draw(lens)
    draw.ellipse((34, 34, 214, 214), outline=GOLD, width=20)
    draw.ellipse((68, 68, 180, 180), outline=(244, 231, 194, 180), width=6)
    draw.rounded_rectangle((176, 170, 286, 204), radius=16, fill=GOLD)
    draw.rounded_rectangle((230, 198, 272, 294), radius=16, fill=(160, 112, 42, 255))
    rotated = lens.rotate(-28, expand=True)
    base.alpha_composite(rotated, (560, 576))


def add_mark(base: Image.Image) -> None:
    seal = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(seal)
    draw.ellipse((160, 164, 284, 288), fill=GOLD)
    draw.ellipse((180, 184, 264, 268), fill=INK)
    draw.rounded_rectangle((210, 204, 234, 248), radius=12, fill=GOLD)
    draw.rounded_rectangle((210, 204, 252, 226), radius=12, fill=GOLD)
    base.alpha_composite(seal)


def build_icon() -> Image.Image:
    icon = build_gradient(SIZE)
    icon = add_glow(icon)
    draw_frame(icon)
    add_documents(icon)
    add_magnifier(icon)
    add_mark(icon)
    return icon


def main() -> int:
    project_root = Path(__file__).resolve().parent
    assets_dir = project_root / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)

    icon = build_icon()
    png_path = assets_dir / "extrator_precatorios.png"
    ico_path = assets_dir / "extrator_precatorios.ico"

    icon.save(png_path)
    sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    icon.save(ico_path, sizes=sizes)

    print(f"PNG criado em: {png_path}")
    print(f"ICO criado em: {ico_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

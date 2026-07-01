#!/usr/bin/env python3
from __future__ import annotations

import os
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent
PNG_DIR = ROOT / "png"
COLOR_DIR = ROOT / "png-color"
ZIP_PATH = ROOT / "slack-text-reactions-20260701-v2.zip"
COLOR_ZIP_PATH = ROOT / "slack-text-reactions-20260701-v2-colors.zip"
PREVIEW_PATH = ROOT / "preview.png"
COLOR_PREVIEW_PATH = ROOT / "preview-colors.png"
FONT_CANDIDATES = [
    Path(os.environ["SLACK_EMOJI_FONT"]) if os.environ.get("SLACK_EMOJI_FONT") else None,
    Path("/home/ubuntu/.local/share/fonts/NotoSansCJKjp-Bold.otf"),
    Path.home() / ".local/share/fonts/NotoSansCJKjp-Bold.otf",
    Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"),
    Path("/usr/share/fonts/opentype/noto/NotoSansCJKjp-Bold.otf"),
]
FONT = next((p for p in FONT_CANDIDATES if p is not None and p.exists()), None)

SIZE = 128
SCALE = 8
W = SIZE * SCALE

# Slack emoji are commonly rendered around 16-24px in message UI.
# Treat each item as a compact logo, not as ordinary text.
ITEMS = [
    ("youkoso", "ようこそ", "#159947"),
    ("kansha", "感謝", "#047857"),
    ("kansha_kangeki", "感謝感激", "#047857"),
    ("sasuga", "さすが", "#B45309"),
    ("sugosugi", "すごすぎ", "#E23A1A"),
    ("tensai", "天才", "#C78A00"),
    ("nice", "ナイス", "#2563EB"),
    ("kakusan", "拡散", "#1D6FCB"),
    ("uwaa", "うわぁ", "#6D5AE6"),
    ("majika", "マジか", "#DB2777"),
    ("saikyou", "最強", "#D9480F"),
    ("komatta", "困った", "#8B5E3C"),
    ("kanben_shite", "勘弁して", "#795548"),
    ("onajiku", "同じく", "#16877B"),
    ("sorena", "それな", "#008C9E"),
    ("doui", "同意", "#246BBA"),
    ("urayama", "うらやま", "#7E57C2"),
    ("yakekuso", "やけくそ", "#C84A1B"),
    ("namusan", "南無三", "#4B5563"),
    ("ouen", "応援", "#0B9954"),
    ("donmai", "どんまい", "#5667E8"),
    ("daishouri", "大勝利", "#D43822"),
    ("iwai", "祝", "#D28A00"),
    ("noroi", "呪", "#5B2A86"),
    ("kami", "神", "#C58B00"),
    ("muri", "無理", "#5B6472"),
    ("tsurai", "つらい", "#6B7280"),
    ("tasukaru", "助かる", "#047857"),
    ("kanmuryou", "感無量", "#0F766E"),
    ("thanks", "サンクス", "#2563EB"),
    ("doumo", "どうも", "#16877B"),
    ("wakaru", "わかる", "#008C9E"),
    ("tashikani", "たしかに", "#246BBA"),
    ("aruaru", "あるある", "#7E57C2"),
    ("kyoukan", "共感", "#0B9954"),
    ("igi_nashi", "異議なし", "#374151"),
    ("ganbare", "がんばれ", "#D9480F"),
    ("odaijini", "お大事に", "#159947"),
    ("muri_sezu", "無理せず", "#5B6472"),
    ("fight", "ファイト", "#E23A1A"),
]

# Slackの小表示で沈みにくい、濃いめの共通カラーバリエーション。
# 黄色系は純黄色ではなくゴールド寄りにして、白背景でも読める濃度にする。
COLOR_VARIANTS = [
    ("red", "#E23A1A"),
    ("orange", "#D9480F"),
    ("green", "#159947"),
    ("blue", "#2563EB"),
    ("purple", "#7E57C2"),
    ("yellow", "#C58B00"),
    ("gray", "#374151"),
]


def _font(px: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(str(FONT), px * SCALE)


def _crop_alpha(img: Image.Image) -> Image.Image:
    bbox = img.getchannel("A").getbbox()
    if bbox is None:
        return img
    # Keep a little transparent breathing room so anti-aliased edge pixels survive cropping.
    pad = 3 * SCALE
    left = max(0, bbox[0] - pad)
    top = max(0, bbox[1] - pad)
    right = min(img.width, bbox[2] + pad)
    bottom = min(img.height, bbox[3] + pad)
    return img.crop((left, top, right, bottom))


def _raw_text(text: str, color: str, font_px: int, stroke_px: int) -> Image.Image:
    font = _font(font_px)
    sw = stroke_px * SCALE
    pad = 96 * SCALE
    # Oversized scratch canvas; cropped by alpha afterwards.
    scratch = Image.new("RGBA", (1400 * SCALE, 420 * SCALE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(scratch)
    # Same-color stroke is used only as synthetic weight. It is not a visible outline.
    draw.text((pad, pad), text, font=font, fill=color, stroke_width=sw, stroke_fill=color)
    return _crop_alpha(scratch)


def _paste_center(canvas: Image.Image, part: Image.Image) -> None:
    x = (canvas.width - part.width) // 2
    y = (canvas.height - part.height) // 2
    canvas.alpha_composite(part, (x, y))


def _resize_exact(part: Image.Image, target_w: int, target_h: int) -> Image.Image:
    return part.resize((target_w * SCALE, target_h * SCALE), Image.Resampling.LANCZOS)


def _line(text: str, color: str) -> Image.Image:
    n = len(text)
    if n == 1:
        target = (124, 124)
        raw = _raw_text(text, color, 124, 2)
    elif n == 2:
        target = (126, 100)
        raw = _raw_text(text, color, 110, 2)
    else:
        # Three-character labels are intentionally horizontally condensed.
        # Uniform scaling makes them too small in Slack.
        target = (127, 84)
        raw = _raw_text(text, color, 112, 2)
    part = _resize_exact(raw, *target)
    canvas = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    _paste_center(canvas, part)
    return canvas.resize((SIZE, SIZE), Image.Resampling.LANCZOS)


def _grid(text: str, color: str) -> Image.Image:
    lines = [text[:2], text[2:4]]
    raw_lines = [_raw_text(line, color, 94, 2) for line in lines]
    gap = -11 * SCALE
    width = max(line.width for line in raw_lines)
    height = raw_lines[0].height + raw_lines[1].height + gap
    composed = Image.new("RGBA", (width + 24 * SCALE, height + 12 * SCALE), (0, 0, 0, 0))
    y = 6 * SCALE
    for line in raw_lines:
        x = (composed.width - line.width) // 2
        composed.alpha_composite(line, (x, y))
        y += line.height + gap
    composed = _crop_alpha(composed)
    part = _resize_exact(composed, 126, 126)
    canvas = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    _paste_center(canvas, part)
    return canvas.resize((SIZE, SIZE), Image.Resampling.LANCZOS)


def make_emoji(text: str, color: str) -> Image.Image:
    if len(text) == 4:
        return _grid(text, color)
    return _line(text, color)


def make_preview(items: list[tuple[str, str, str]]) -> Image.Image:
    cols = 7
    cell = 128
    rows = (len(items) + cols - 1) // cols
    bg = Image.new("RGBA", (cols * cell, rows * cell), "#f8f8f8")
    for idx, (slug, text, color) in enumerate(items):
        img = Image.open(PNG_DIR / f"{slug}.png").convert("RGBA")
        x = (idx % cols) * cell
        y = (idx // cols) * cell
        bg.alpha_composite(img, (x, y))
    return bg.convert("RGB")


def make_color_preview(items: list[tuple[str, str, str]]) -> Image.Image:
    cols = len(COLOR_VARIANTS)
    cell = 128
    rows = len(items)
    bg = Image.new("RGBA", (cols * cell, rows * cell), "#f8f8f8")
    for row, (slug, _text, _color) in enumerate(items):
        for col, (variant, _variant_color) in enumerate(COLOR_VARIANTS):
            img = Image.open(COLOR_DIR / f"{slug}_{variant}.png").convert("RGBA")
            x = col * cell
            y = row * cell
            bg.alpha_composite(img, (x, y))
    return bg.convert("RGB")


def main() -> None:
    if FONT is None:
        raise SystemExit("font not found: set SLACK_EMOJI_FONT to a Japanese gothic font")
    PNG_DIR.mkdir(parents=True, exist_ok=True)
    COLOR_DIR.mkdir(parents=True, exist_ok=True)
    for slug, text, color in ITEMS:
        img = make_emoji(text, color)
        img.save(PNG_DIR / f"{slug}.png", optimize=True)
        for variant, variant_color in COLOR_VARIANTS:
            variant_img = make_emoji(text, variant_color)
            variant_img.save(COLOR_DIR / f"{slug}_{variant}.png", optimize=True)
    preview = make_preview(ITEMS)
    preview.save(PREVIEW_PATH, optimize=True)
    color_preview = make_color_preview(ITEMS)
    color_preview.save(COLOR_PREVIEW_PATH, optimize=True)
    with ZipFile(ZIP_PATH, "w", compression=ZIP_DEFLATED, compresslevel=9) as zf:
        for slug, _, _ in ITEMS:
            path = PNG_DIR / f"{slug}.png"
            zf.write(path, arcname=f"png/{path.name}")
    with ZipFile(COLOR_ZIP_PATH, "w", compression=ZIP_DEFLATED, compresslevel=9) as zf:
        for slug, _, _ in ITEMS:
            for variant, _ in COLOR_VARIANTS:
                path = COLOR_DIR / f"{slug}_{variant}.png"
                zf.write(path, arcname=f"png-color/{path.name}")


if __name__ == "__main__":
    main()

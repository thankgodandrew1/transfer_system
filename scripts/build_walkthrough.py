"""Build the sanitized, captioned walkthrough video used by the Guide page."""
from __future__ import annotations

import textwrap
from pathlib import Path

import imageio.v2 as imageio
import numpy as np
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
WIDTH, HEIGHT = 1280, 720
FPS = 12
SECONDS_PER_SLIDE = 7
FRAMES_PER_SLIDE = FPS * SECONDS_PER_SLIDE
GREEN = "#123f2a"
GREEN_DARK = "#0b2b1d"
GREEN_LIGHT = "#e4f2eb"
GOLD = "#c9a227"
GOLD_LIGHT = "#fbf5df"
INK = "#17211b"
MUTED = "#66736b"
PAPER = "#ffffff"
CANVAS = "#f5f7f5"


def font(size: int, bold: bool = False, serif: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = []
    if serif:
        candidates.extend([
            Path("C:/Windows/Fonts/georgiab.ttf") if bold else Path("C:/Windows/Fonts/georgia.ttf"),
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf") if bold else Path("/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf"),
        ])
    else:
        candidates.extend([
            Path("C:/Windows/Fonts/segoeuib.ttf") if bold else Path("C:/Windows/Fonts/segoeui.ttf"),
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf") if bold else Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        ])
    for candidate in candidates:
        if candidate.exists():
            return ImageFont.truetype(str(candidate), size)
    return ImageFont.load_default()


def rounded(draw: ImageDraw.ImageDraw, box, fill, outline=None, radius=22, width=2):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def text(draw: ImageDraw.ImageDraw, xy, value: str, size: int, fill=INK, bold=False, serif=False, anchor=None):
    draw.text(xy, value, font=font(size, bold=bold, serif=serif), fill=fill, anchor=anchor)


def wrapped(draw: ImageDraw.ImageDraw, xy, value: str, width: int, size: int, fill=MUTED, bold=False, serif=False, spacing=10):
    lines = textwrap.wrap(value, width=width)
    draw.multiline_text(xy, "\n".join(lines), font=font(size, bold=bold, serif=serif), fill=fill, spacing=spacing)


def base(step: str, title_value: str, subtitle: str, subtitle_width: int = 70) -> tuple[Image.Image, ImageDraw.ImageDraw]:
    image = Image.new("RGB", (WIDTH, HEIGHT), CANVAS)
    draw = ImageDraw.Draw(image)
    draw.rectangle((0, 0, WIDTH, 82), fill=GREEN_DARK)
    rounded(draw, (42, 17, 94, 67), GREEN, GOLD, radius=14, width=2)
    text(draw, (68, 42), "TN", 17, GOLD, bold=True, anchor="mm")
    text(draw, (112, 27), "TRANSFER NEWS GENERATOR", 19, PAPER, bold=True)
    text(draw, (112, 52), "SANITIZED USER WALKTHROUGH", 11, "#b9c9c0", bold=True)
    text(draw, (70, 122), step, 15, GOLD, bold=True)
    title_lines = title_value.splitlines()
    draw.multiline_text(
        (70, 155), title_value, font=font(42, bold=True, serif=True),
        fill=GREEN_DARK, spacing=0,
    )
    wrapped(draw, (72, 216 + (len(title_lines) - 1) * 48), subtitle, subtitle_width, 22, MUTED, spacing=7)
    return image, draw


def slide_welcome() -> Image.Image:
    image, draw = base(
        "WELCOME", "Verified Transfer News,\nin one complete run.",
        "Use the browser from any device to turn four transfer exports into a reviewable publishing package.",
        subtitle_width=45,
    )
    source = ROOT / "tmp" / "browser" / "home-desktop.png"
    if source.exists():
        screen = Image.open(source).convert("RGB")
        screen = screen.crop((0, 0, screen.width, min(screen.height, 385)))
        screen.thumbnail((610, 365))
        rounded(draw, (620, 120, 1215, 650), PAPER, "#dce5df", radius=24)
        image.paste(screen, (642, 151))
    rounded(draw, (72, 352, 520, 427), GREEN, radius=14)
    text(draw, (296, 389), "GENERATE TRANSFER NEWS", 18, PAPER, bold=True, anchor="mm")
    text(draw, (72, 482), "What you get", 16, GREEN, bold=True)
    text(draw, (72, 522), "PDF  •  Word  •  Excel  •  CSV  •  ZIP", 23, INK, bold=True)
    text(draw, (72, 568), "Plus an updated roster and statistics", 18, MUTED)
    return image


def slide_uploads() -> Image.Image:
    image, draw = base("STEP 1", "Upload the four source records.", "The two PDFs provide the transfer layout and previous news. The two Excel reports provide authoritative current and prior roster facts.")
    cards = [
        ("TM", "Current Transfer Management", "PDF from iMOS"),
        ("PN", "Previous Transfer News", "Last updated-roster PDF"),
        ("CR", "Current Transfer Report", "Current roster XLSX"),
        ("OR", "Previous Transfer Report", "Prior roster XLSX"),
    ]
    for index, (code, label, detail) in enumerate(cards):
        col, row = index % 2, index // 2
        x, y = 70 + col * 575, 348 + row * 125
        rounded(draw, (x, y, x + 535, y + 100), PAPER, "#cbd7cf", radius=16)
        rounded(draw, (x + 18, y + 20, x + 78, y + 80), GREEN_LIGHT, radius=12)
        text(draw, (x + 48, y + 50), code, 16, GREEN, bold=True, anchor="mm")
        text(draw, (x + 98, y + 24), label, 18, INK, bold=True)
        text(draw, (x + 98, y + 55), detail, 15, MUTED)
    text(draw, (70, 622), "Tip: use the updated roster produced by the previous run as the next cycle’s Previous Transfer News.", 16, GREEN, bold=True)
    return image


def slide_details() -> Image.Image:
    image, draw = base("STEP 2", "Review the publication details.", "Confirm the mission, president, preparer, and detected transfer title. The title override is optional.")
    fields = [
        ("Mission name", "DEMO MISSION"),
        ("Mission president", "PRESIDENT EXAMPLE"),
        ("Prepared by", "TRANSFER TEAM"),
        ("Transfer title", "SEPTEMBER 2026 TRANSFER NEWS"),
    ]
    for index, (label, value) in enumerate(fields):
        col, row = index % 2, index // 2
        x, y = 70 + col * 575, 346 + row * 132
        text(draw, (x, y), label.upper(), 13, GREEN, bold=True)
        rounded(draw, (x, y + 28, x + 525, y + 91), PAPER, "#cbd7cf", radius=10)
        text(draw, (x + 18, y + 49), value, 17, INK, bold=True)
    rounded(draw, (70, 616, 1190, 671), GOLD_LIGHT, radius=12)
    text(draw, (630, 644), "Never use real names or confidential files in a public demo or support request.", 17, "#7a6010", bold=True, anchor="mm")
    return image


def slide_options() -> Image.Image:
    image, draw = base("STEP 3", "Choose optional outputs and controls.", "Keep the updated roster and statistics selected for a complete package. Upload a template or manual-corrections CSV only when needed.")
    items = [
        ("ON", "Updated roster", "Becomes next cycle’s previous-news input", GREEN_LIGHT),
        ("ON", "Statistics", "Word and PDF summaries", GREEN_LIGHT),
        ("+", "Word template", "Optional custom DOCX layout", GOLD_LIGHT),
        ("+", "Manual corrections", "CSV columns: Name, Field, Value", GOLD_LIGHT),
    ]
    for index, (mark, label, detail, fill) in enumerate(items):
        x, y = 70 + index * 285, 365
        rounded(draw, (x, y, x + 255, y + 210), fill, "#dce5df", radius=18)
        rounded(draw, (x + 20, y + 20, x + 72, y + 72), GREEN, radius=14)
        text(draw, (x + 46, y + 46), mark, 17 if mark == "ON" else 24, PAPER, bold=True, anchor="mm")
        text(draw, (x + 20, y + 100), label, 20, INK, bold=True)
        wrapped(draw, (x + 20, y + 136), detail, 22, 15, MUTED, spacing=5)
    text(draw, (70, 625), "Recommended default: keep both checkboxes selected.", 18, GREEN, bold=True)
    return image


def slide_generate() -> Image.Image:
    image, draw = base("STEP 4", "Generate and keep the tab open.", "The service validates files, matches missionaries, applies ground-truth verification, and renders the package.")
    rounded(draw, (155, 340, 1125, 615), PAPER, "#dce5df", radius=26)
    stages = ["Validate uploads", "Match assignments", "Verify reports", "Render outputs", "Build ZIP"]
    for index, label in enumerate(stages):
        x = 235 + index * 180
        draw.line((x, 445, x + 135, 445), fill="#b9c9c0", width=5) if index < len(stages) - 1 else None
        draw.ellipse((x - 24, 421, x + 24, 469), fill=GREEN if index < 4 else GOLD)
        text(draw, (x, 445), "OK" if index < 4 else "...", 13, PAPER, bold=True, anchor="mm")
        text(draw, (x, 500), label, 15, INK, bold=True, anchor="ma")
    rounded(draw, (260, 555, 1020, 566), GREEN_LIGHT, radius=6)
    rounded(draw, (260, 555, 895, 566), GREEN, radius=6)
    return image


def slide_review() -> Image.Image:
    image, draw = base("STEP 5", "Review before publishing.", "Download the ZIP, then inspect the Verification and Changes sheets. If the run is blocked, fix every BLOCKING row and run again.")
    metrics = [("96", "MISSIONARIES"), ("8", "ZONES"), ("4", "NEW"), ("0", "BLOCKING")]
    for index, (value, label) in enumerate(metrics):
        x = 70 + index * 285
        rounded(draw, (x, 345, x + 255, 465), PAPER, "#dce5df", radius=16)
        text(draw, (x + 127, 390), value, 36, GREEN, bold=True, anchor="mm")
        text(draw, (x + 127, 433), label, 13, MUTED, bold=True, anchor="mm")
    downloads = ["Complete ZIP", "Transfer News PDF", "Review workbook", "Updated roster", "Statistics"]
    for index, label in enumerate(downloads):
        x = 70 + (index % 3) * 380
        y = 510 + (index // 3) * 66
        rounded(draw, (x, y, x + 350, y + 48), GREEN_LIGHT, radius=10)
        text(draw, (x + 18, y + 14), "↓", 18, GREEN, bold=True)
        text(draw, (x + 52, y + 15), label, 16, INK, bold=True)
    return image


def slide_privacy() -> Image.Image:
    image, draw = base("FINISH", "Download promptly. Protect the records.", "Each job is isolated and expires automatically. Use the shared access key on public hosting and keep all real transfer files out of GitHub.")
    points = [
        "Download the package before the expiry time.",
        "Keep real PDFs, spreadsheets, and output files private.",
        "Store only code, templates, and sanitized examples in GitHub.",
        "Use the written guide whenever you need the full checklist.",
    ]
    for index, point in enumerate(points):
        y = 348 + index * 70
        draw.ellipse((78, y + 2, 104, y + 28), fill=GOLD)
        text(draw, (91, y + 15), str(index + 1), 14, GREEN_DARK, bold=True, anchor="mm")
        text(draw, (125, y), point, 21, INK, bold=True)
    rounded(draw, (760, 346, 1170, 625), GREEN_DARK, radius=24)
    text(draw, (965, 410), "READY", 15, GOLD, bold=True, anchor="mm")
    text(draw, (965, 472), "Generate.", 34, PAPER, bold=True, serif=True, anchor="mm")
    text(draw, (965, 520), "Review.", 34, PAPER, bold=True, serif=True, anchor="mm")
    text(draw, (965, 568), "Publish safely.", 34, PAPER, bold=True, serif=True, anchor="mm")
    return image


def build_video() -> None:
    slides = [slide_welcome(), slide_uploads(), slide_details(), slide_options(), slide_generate(), slide_review(), slide_privacy()]
    output = ROOT / "static" / "walkthrough.mp4"
    poster = ROOT / "static" / "video-poster.png"
    slides[0].save(poster, optimize=True)
    writer = imageio.get_writer(
        output,
        fps=FPS,
        codec="libx264",
        quality=8,
        pixelformat="yuv420p",
        ffmpeg_log_level="error",
    )
    try:
        for slide_index, slide in enumerate(slides):
            current = np.asarray(slide)
            previous = np.asarray(slides[slide_index - 1]) if slide_index else current
            for frame_index in range(FRAMES_PER_SLIDE):
                if slide_index and frame_index < FPS // 2:
                    alpha = frame_index / max(FPS // 2 - 1, 1)
                    frame = (previous * (1 - alpha) + current * alpha).astype(np.uint8)
                else:
                    frame = current
                writer.append_data(frame)
    finally:
        writer.close()
    print(f"Created {output} and {poster}")


if __name__ == "__main__":
    build_video()

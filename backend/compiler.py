import base64
import io
import os
import re
import urllib.parse

import requests
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

try:
    from PIL import Image
except ImportError:  # pragma: no cover
    Image = None


SLIDE_WIDTH = 13.333
SLIDE_HEIGHT = 7.5

PALETTE = {
    "navy": RGBColor(15, 23, 42),
    "ink": RGBColor(30, 41, 59),
    "slate": RGBColor(71, 85, 105),
    "muted": RGBColor(100, 116, 139),
    "line": RGBColor(203, 213, 225),
    "panel": RGBColor(248, 250, 252),
    "soft_blue": RGBColor(219, 234, 254),
    "accent": RGBColor(14, 116, 144),
    "accent_dark": RGBColor(8, 47, 73),
    "white": RGBColor(255, 255, 255),
}

IMAGE_STYLE_SUFFIX = (
    "premium editorial presentation art, abstract 3d illustration, "
    "clean geometry, cohesive corporate theme, vibrant but elegant, "
    "not photorealistic, high contrast lighting"
)


def clean_base64_string(b64_string: str) -> str:
    clean = str(b64_string or "").strip()
    if "base64," in clean:
        clean = clean.split("base64,", 1)[1]
    clean = clean.strip().strip("()").strip("'").strip('"').strip()
    clean = re.sub(r"\s+", "", clean)
    clean += "=" * ((4 - len(clean) % 4) % 4)
    return clean


def find_best_match(topic_ref, harvested_keys):
    if not topic_ref:
        return None
    if topic_ref in harvested_keys:
        return topic_ref

    normalized_ref = _normalize_key(topic_ref)
    for key in harvested_keys:
        normalized_key = _normalize_key(key)
        if normalized_ref == normalized_key:
            return key
        if normalized_ref in normalized_key or normalized_key in normalized_ref:
            return key
    return None


def build_powerpoint(master_blueprint_json: list, harvested_images: dict, output_filename="Final_Deck.pptx"):
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_WIDTH)
    prs.slide_height = Inches(SLIDE_HEIGHT)
    blank_slide_layout = prs.slide_layouts[6]

    harvested_images = harvested_images or {}

    for slide_index, slide_data in enumerate(master_blueprint_json or [], start=1):
        slide = prs.slides.add_slide(blank_slide_layout)
        slide_data = slide_data or {}
        layout_type = slide_data.get("layout_type", "split_left_image_right_cards")

        _paint_base_background(slide)

        background_stream = _resolve_background_stream(slide_data)
        if background_stream:
            _add_picture_cover(
                slide,
                background_stream,
                Inches(0),
                Inches(0),
                Inches(SLIDE_WIDTH),
                Inches(SLIDE_HEIGHT),
            )
            _add_overlay_wash(slide, 0.18 if layout_type == "hero_background_with_floating_cards" else 0.08)

        _add_decorative_accents(slide, layout_type, has_background=bool(background_stream))
        _draw_title(slide, slide_data.get("title", f"Slide {slide_index}"), bool(background_stream))

        image_stream = _resolve_foreground_stream(slide_data, harvested_images)
        blocks = _normalize_blocks(slide_data)

        _render_layout(slide, layout_type, blocks, image_stream, bool(background_stream))

    prs.save(output_filename)
    return output_filename


def _normalize_key(value):
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


def _resolve_background_stream(slide_data):
    bg_info = slide_data.get("background", {}) or {}
    if not bg_info.get("has_background_image"):
        return None

    prompt = bg_info.get("background_image_prompt")
    if not prompt:
        return None

    return _generate_image_from_prompt(
        f"{prompt}, wide cinematic backdrop, soft gradients, layered composition",
        width=1280,
        height=720,
        timeout=45,
    )


def _resolve_foreground_stream(slide_data, harvested_images):
    existing_chart_info = slide_data.get("existing_chart", {}) or {}
    topic_ref = existing_chart_info.get("chart_topic_reference")
    best_match_key = find_best_match(topic_ref, list(harvested_images.keys()))

    if existing_chart_info.get("has_existing_chart") and best_match_key:
        stream = _first_image_stream(harvested_images.get(best_match_key))
        if stream:
            print(f"Using extracted chart for topic: {best_match_key}")
            return stream

    foreground_info = slide_data.get("foreground_image", {}) or {}
    if foreground_info.get("has_foreground_image"):
        prompt = foreground_info.get("foreground_image_prompt")
        if prompt:
            return _generate_image_from_prompt(
                f"{prompt}, isolated subject, transparent style, presentation-ready asset",
                width=1024,
                height=768,
                timeout=45,
            )

    fallback_candidates = []
    if topic_ref and best_match_key:
        fallback_candidates.append(harvested_images.get(best_match_key))
    for candidate in fallback_candidates:
        stream = _first_image_stream(candidate)
        if stream:
            return stream

    return None


def _generate_image_from_prompt(prompt, width, height, timeout):
    styled_prompt = f"{prompt}, {IMAGE_STYLE_SUFFIX}"
    encoded_prompt = urllib.parse.quote(styled_prompt)
    image_url = (
        f"https://image.pollinations.ai/prompt/{encoded_prompt}"
        f"?width={width}&height={height}&nologo=true"
    )

    try:
        response = requests.get(image_url, timeout=timeout)
        response.raise_for_status()
        return io.BytesIO(response.content)
    except Exception as exc:
        print(f"Image generation skipped: {exc}")
        return None


def _first_image_stream(raw_value):
    if raw_value is None:
        return None

    if isinstance(raw_value, (list, tuple)):
        for item in raw_value:
            stream = _coerce_image_stream(item)
            if stream:
                return stream
        return None

    return _coerce_image_stream(raw_value)


def _coerce_image_stream(image_value):
    if image_value is None:
        return None

    if isinstance(image_value, io.BytesIO):
        image_value.seek(0)
        return io.BytesIO(image_value.read())

    if isinstance(image_value, bytes):
        return io.BytesIO(image_value)

    if isinstance(image_value, dict):
        for key in ("data", "base64", "image_base64", "content", "path", "url"):
            if image_value.get(key):
                return _coerce_image_stream(image_value[key])
        return None

    if not isinstance(image_value, str):
        return None

    candidate = image_value.strip()
    if not candidate:
        return None

    if candidate.startswith("(") and candidate.endswith(")"):
        candidate = candidate[1:-1].strip()

    if candidate.startswith("data:image/") or "base64," in candidate:
        try:
            return io.BytesIO(base64.b64decode(clean_base64_string(candidate)))
        except Exception as exc:
            print(f"Failed to decode base64 image: {exc}")
            return None

    if os.path.exists(candidate):
        try:
            with open(candidate, "rb") as image_file:
                return io.BytesIO(image_file.read())
        except OSError as exc:
            print(f"Failed to read image file {candidate}: {exc}")
            return None

    if candidate.startswith("http://") or candidate.startswith("https://"):
        try:
            response = requests.get(candidate, timeout=30)
            response.raise_for_status()
            return io.BytesIO(response.content)
        except Exception as exc:
            print(f"Failed to fetch image URL: {exc}")
            return None

    try:
        return io.BytesIO(base64.b64decode(clean_base64_string(candidate)))
    except Exception:
        return None


def _normalize_blocks(slide_data):
    blocks = slide_data.get("content_blocks")
    if isinstance(blocks, list) and blocks:
        normalized = []
        for block in blocks:
            if not isinstance(block, dict):
                continue

            bullets = block.get("bullets", [])
            if isinstance(bullets, str):
                bullets = [bullets]
            bullets = [str(item).strip() for item in bullets if str(item).strip()]

            body = block.get("body") or block.get("text") or block.get("description") or ""
            normalized.append(
                {
                    "block_title": str(block.get("block_title") or block.get("title") or "").strip(),
                    "bullets": bullets,
                    "body": str(body).strip(),
                }
            )

        if normalized:
            return normalized

    content = slide_data.get("content", [])
    if isinstance(content, list) and content:
        bullets = [item.get("text", "").strip() for item in content if isinstance(item, dict) and item.get("type") == "bullet"]
        paragraphs = [item.get("text", "").strip() for item in content if isinstance(item, dict) and item.get("type") != "bullet"]
        return [
            {
                "block_title": "Overview",
                "bullets": [item for item in bullets if item],
                "body": " ".join(item for item in paragraphs if item),
            }
        ]

    raw_text = str(slide_data.get("raw_text") or "").strip()
    if raw_text:
        lines = [line.strip("-* ").strip() for line in raw_text.splitlines() if line.strip()]
        return [{"block_title": "Overview", "bullets": lines[:5], "body": ""}]

    return [{"block_title": "Key Takeaway", "bullets": [], "body": ""}]


def _render_layout(slide, layout_type, blocks, image_stream, has_background):
    if layout_type == "top_image_bottom_cards":
        _layout_top_image_bottom_cards(slide, blocks, image_stream)
    elif layout_type == "hero_background_with_floating_cards":
        _layout_hero_background_with_floating_cards(slide, blocks, image_stream, has_background)
    elif layout_type == "multi_column_cards":
        _layout_multi_column_cards(slide, blocks, image_stream, has_background)
    elif layout_type == "diagonal_text_split":
        _layout_diagonal_text_split(slide, blocks, image_stream)
    else:
        _layout_split_left_image_right_cards(slide, blocks, image_stream)


def _layout_top_image_bottom_cards(slide, blocks, image_stream):
    if image_stream:
        _add_picture_contain(slide, image_stream, Inches(0.75), Inches(1.3), Inches(11.85), Inches(2.6))
    else:
        _add_placeholder_panel(slide, Inches(0.95), Inches(1.45), Inches(11.45), Inches(2.25), "Visual summary")

    for index, (left, top, width, height) in enumerate(_horizontal_card_specs(len(blocks), Inches(4.25), Inches(2.5))):
        _add_card(slide, left, top, width, height, blocks[index], dark=False)


def _layout_hero_background_with_floating_cards(slide, blocks, image_stream, has_background):
    if not has_background:
        _add_overlay_wash(slide, 0.04)

    if image_stream:
        _add_picture_contain(slide, image_stream, Inches(7.8), Inches(1.25), Inches(4.6), Inches(2.5))

    specs = [
        (Inches(0.75), Inches(1.45), Inches(4.4), Inches(2.0)),
        (Inches(1.05), Inches(4.05), Inches(3.7), Inches(2.1)),
        (Inches(5.2), Inches(4.35), Inches(3.3), Inches(1.8)),
        (Inches(8.95), Inches(4.0), Inches(3.35), Inches(2.0)),
    ]
    for index, block in enumerate(blocks[: len(specs)]):
        left, top, width, height = specs[index]
        _add_card(slide, left, top, width, height, block, dark=True)


def _layout_multi_column_cards(slide, blocks, image_stream, has_background):
    if image_stream:
        _add_picture_contain(slide, image_stream, Inches(9.9), Inches(0.95), Inches(2.5), Inches(1.5))

    if not has_background:
        _add_band(slide, Inches(0.6), Inches(1.4), Inches(12.1), Inches(4.9), PALETTE["soft_blue"], transparency=0.35)

    for index, (left, top, width, height) in enumerate(_horizontal_card_specs(len(blocks), Inches(1.75), Inches(4.35), gutter=0.22)):
        _add_card(slide, left, top, width, height, blocks[index], dark=False)


def _layout_diagonal_text_split(slide, blocks, image_stream):
    if blocks:
        _add_card(slide, Inches(0.7), Inches(1.45), Inches(4.15), Inches(2.25), blocks[0], dark=False)
    if len(blocks) > 1:
        _add_card(slide, Inches(8.35), Inches(4.0), Inches(4.2), Inches(2.3), blocks[1], dark=False)
    if len(blocks) > 2:
        _add_card(slide, Inches(0.95), Inches(4.25), Inches(3.5), Inches(1.55), blocks[2], dark=True)

    if image_stream:
        _add_picture_contain(slide, image_stream, Inches(4.55), Inches(2.0), Inches(4.2), Inches(3.05))
    else:
        _add_placeholder_panel(slide, Inches(4.85), Inches(2.35), Inches(3.65), Inches(2.35), "Central visual")


def _layout_split_left_image_right_cards(slide, blocks, image_stream):
    if image_stream:
        _add_picture_contain(slide, image_stream, Inches(0.7), Inches(1.45), Inches(6.05), Inches(4.9))
    else:
        _add_placeholder_panel(slide, Inches(0.9), Inches(1.75), Inches(5.6), Inches(4.3), "Primary visual")

    max_cards = min(len(blocks), 3)
    card_height = 4.9 / max(max_cards, 1)
    for index in range(max_cards):
        _add_card(
            slide,
            Inches(7.1),
            Inches(1.45 + (index * card_height)),
            Inches(5.45),
            Inches(card_height - 0.15),
            blocks[index],
            dark=False,
        )


def _horizontal_card_specs(count, top, height, gutter=0.28):
    count = max(1, min(count, 4))
    total_width = 12.1
    usable_width = total_width - (gutter * (count - 1))
    card_width = usable_width / count
    specs = []
    current_left = 0.6
    for _ in range(count):
        specs.append((Inches(current_left), top, Inches(card_width), height))
        current_left += card_width + gutter
    return specs


def _paint_base_background(slide):
    base = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(0),
        Inches(0),
        Inches(SLIDE_WIDTH),
        Inches(SLIDE_HEIGHT),
    )
    base.fill.solid()
    base.fill.fore_color.rgb = RGBColor(242, 247, 251)
    base.line.fill.background()

    glow = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.OVAL,
        Inches(-0.9),
        Inches(-1.2),
        Inches(4.2),
        Inches(4.2),
    )
    glow.fill.solid()
    glow.fill.fore_color.rgb = RGBColor(186, 230, 253)
    glow.fill.transparency = 0.45
    glow.line.fill.background()

    corner = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.OVAL,
        Inches(10.6),
        Inches(5.2),
        Inches(3.2),
        Inches(3.2),
    )
    corner.fill.solid()
    corner.fill.fore_color.rgb = RGBColor(191, 219, 254)
    corner.fill.transparency = 0.58
    corner.line.fill.background()


def _add_decorative_accents(slide, layout_type, has_background):
    accent = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(0.55),
        Inches(0.62),
        Inches(1.15),
        Inches(0.1),
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = PALETTE["accent"]
    accent.line.fill.background()

    if layout_type in {"split_left_image_right_cards", "diagonal_text_split"}:
        side_bar = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.RECTANGLE,
            Inches(6.74),
            Inches(1.35),
            Inches(0.08),
            Inches(5.1),
        )
        side_bar.fill.solid()
        side_bar.fill.fore_color.rgb = PALETTE["line"]
        side_bar.fill.transparency = 0.25 if has_background else 0.0
        side_bar.line.fill.background()


def _add_overlay_wash(slide, transparency):
    overlay = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(0),
        Inches(0),
        Inches(SLIDE_WIDTH),
        Inches(SLIDE_HEIGHT),
    )
    overlay.fill.solid()
    overlay.fill.fore_color.rgb = PALETTE["white"]
    overlay.fill.transparency = transparency
    overlay.line.fill.background()


def _add_band(slide, left, top, width, height, color, transparency=0.0):
    band = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, left, top, width, height)
    band.fill.solid()
    band.fill.fore_color.rgb = color
    band.fill.transparency = transparency
    band.line.fill.background()
    return band


def _draw_title(slide, title_text, on_busy_background):
    eyebrow = slide.shapes.add_textbox(Inches(0.75), Inches(0.55), Inches(3.0), Inches(0.3))
    eyebrow_tf = eyebrow.text_frame
    eyebrow_tf.clear()
    eyebrow_p = eyebrow_tf.paragraphs[0]
    eyebrow_p.text = "AI PREPARED DECK"
    eyebrow_p.font.size = Pt(10)
    eyebrow_p.font.bold = True
    eyebrow_p.font.color.rgb = PALETTE["accent_dark"] if not on_busy_background else PALETTE["accent"]
    eyebrow_p.alignment = PP_ALIGN.LEFT

    title_box = slide.shapes.add_textbox(Inches(0.75), Inches(0.78), Inches(11.7), Inches(0.9))
    title_box.text_frame.word_wrap = True
    title_box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = title_box.text_frame.paragraphs[0]
    p.text = str(title_text or "Untitled Slide")
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = PALETTE["navy"] if not on_busy_background else PALETTE["white"]
    p.alignment = PP_ALIGN.LEFT


def _add_card(slide, left, top, width, height, block_data, dark=False):
    card = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = PALETTE["accent_dark"] if dark else PALETTE["white"]
    card.fill.transparency = 0.1 if dark else 0.02
    card.line.color.rgb = RGBColor(125, 211, 252) if dark else PALETTE["line"]
    card.line.width = Pt(1.2)
    _draw_text_block(card, block_data, dark=dark)
    return card


def _add_placeholder_panel(slide, left, top, width, height, label):
    panel = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, left, top, width, height)
    panel.fill.solid()
    panel.fill.fore_color.rgb = RGBColor(226, 232, 240)
    panel.line.color.rgb = PALETTE["line"]
    panel.line.width = Pt(1)
    tf = panel.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = label
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = PALETTE["muted"]
    p.alignment = PP_ALIGN.CENTER
    return panel


def _draw_text_block(shape, block_data, dark=False):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.22)
    tf.margin_right = Inches(0.22)
    tf.margin_top = Inches(0.16)
    tf.margin_bottom = Inches(0.14)
    tf.vertical_anchor = MSO_ANCHOR.TOP

    title = str(block_data.get("block_title") or "Insight").strip()
    title_para = tf.paragraphs[0]
    title_para.text = title
    title_para.font.size = Pt(18 if dark else 17)
    title_para.font.bold = True
    title_para.font.color.rgb = PALETTE["white"] if dark else PALETTE["navy"]
    title_para.space_after = Pt(8)

    body = str(block_data.get("body") or "").strip()
    if body:
        body_para = tf.add_paragraph()
        body_para.text = body
        body_para.font.size = Pt(11.5)
        body_para.font.color.rgb = RGBColor(226, 232, 240) if dark else PALETTE["slate"]
        body_para.space_after = Pt(8)

    bullets = block_data.get("bullets", []) or []
    for bullet in bullets[:5]:
        bullet_para = tf.add_paragraph()
        bullet_para.text = f"- {str(bullet)}"
        bullet_para.level = 0
        bullet_para.font.size = Pt(11.5)
        bullet_para.font.color.rgb = RGBColor(226, 232, 240) if dark else PALETTE["ink"]
        bullet_para.space_after = Pt(4)


def _add_picture_cover(slide, image_stream, left, top, width, height):
    return _add_picture(slide, image_stream, left, top, width, height, mode="cover")


def _add_picture_contain(slide, image_stream, left, top, width, height):
    return _add_picture(slide, image_stream, left, top, width, height, mode="contain")


def _add_picture(slide, image_stream, left, top, width, height, mode):
    if not image_stream:
        return None

    image_bytes = _read_stream_bytes(image_stream)
    if not image_bytes:
        return None

    image_size = _read_image_size(image_bytes)
    if not image_size:
        try:
            return slide.shapes.add_picture(io.BytesIO(image_bytes), left, top, width=width, height=height)
        except Exception as exc:
            print(f"Failed to place image: {exc}")
            return None

    img_width_px, img_height_px = image_size
    if img_width_px <= 0 or img_height_px <= 0:
        return None

    frame_ratio = width / height
    image_ratio = img_width_px / img_height_px

    if mode == "cover":
        if image_ratio > frame_ratio:
            draw_height = height
            draw_width = height * image_ratio
        else:
            draw_width = width
            draw_height = width / image_ratio
    else:
        if image_ratio > frame_ratio:
            draw_width = width
            draw_height = width / image_ratio
        else:
            draw_height = height
            draw_width = height * image_ratio

    draw_left = left + (width - draw_width) / 2
    draw_top = top + (height - draw_height) / 2

    try:
        return slide.shapes.add_picture(io.BytesIO(image_bytes), draw_left, draw_top, width=draw_width, height=draw_height)
    except Exception as exc:
        print(f"Failed to place image: {exc}")
        return None


def _read_stream_bytes(stream):
    if isinstance(stream, io.BytesIO):
        stream.seek(0)
        return stream.read()
    if isinstance(stream, bytes):
        return stream
    return None


def _read_image_size(image_bytes):
    if Image is None:
        return None

    try:
        with Image.open(io.BytesIO(image_bytes)) as image:
            return image.size
    except Exception:
        return None

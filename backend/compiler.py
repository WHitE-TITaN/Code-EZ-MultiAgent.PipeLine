import base64
import hashlib
import io
import os
import random
import re
import time
import urllib.parse

import requests
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Inches, Pt

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:  # pragma: no cover
    Image = None
    ImageDraw = None
    ImageFont = None


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

THEMES = {
    "ocean": {
        "base": RGBColor(242, 247, 251),
        "glow": RGBColor(186, 230, 253),
        "corner": RGBColor(191, 219, 254),
        "accent": RGBColor(14, 116, 144),
        "accent_dark": RGBColor(8, 47, 73),
        "band": RGBColor(219, 234, 254),
        "card_fill": RGBColor(255, 255, 255),
    },
    "midnight": {
        "base": RGBColor(239, 244, 255),
        "glow": RGBColor(165, 180, 252),
        "corner": RGBColor(196, 181, 253),
        "accent": RGBColor(79, 70, 229),
        "accent_dark": RGBColor(49, 46, 129),
        "band": RGBColor(224, 231, 255),
        "card_fill": RGBColor(255, 255, 255),
    },
    "emerald": {
        "base": RGBColor(240, 253, 250),
        "glow": RGBColor(153, 246, 228),
        "corner": RGBColor(187, 247, 208),
        "accent": RGBColor(5, 150, 105),
        "accent_dark": RGBColor(6, 78, 59),
        "band": RGBColor(209, 250, 229),
        "card_fill": RGBColor(255, 255, 255),
    },
    "sunset": {
        "base": RGBColor(255, 247, 237),
        "glow": RGBColor(253, 230, 138),
        "corner": RGBColor(253, 186, 116),
        "accent": RGBColor(234, 88, 12),
        "accent_dark": RGBColor(124, 45, 18),
        "band": RGBColor(254, 215, 170),
        "card_fill": RGBColor(255, 255, 255),
    },
}

IMAGE_STYLE_SUFFIX = (
    "premium editorial presentation art, abstract 3d illustration, "
    "clean geometry, cohesive corporate theme, vibrant but elegant, "
    "not photorealistic, high contrast lighting"
)

HTTP_SESSION = requests.Session()
HTTP_SESSION.headers.update(
    {
        "User-Agent": "CodeEZ-Pipeline/1.0",
        "Accept": "image/*,*/*;q=0.8",
    }
)
IMAGE_CACHE = {}
REMOTE_IMAGE_DISABLED = False


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
    deck_title = _resolve_deck_title(master_blueprint_json, output_filename)
    _add_cover_slide(prs, deck_title)

    for slide_index, slide_data in enumerate(master_blueprint_json or [], start=1):
        slide = prs.slides.add_slide(blank_slide_layout)
        slide_data = slide_data or {}
        theme = _resolve_theme(slide_data, slide_index)
        layout_type = slide_data.get("layout_type", "split_left_image_right_cards")

        _paint_base_background(slide, slide_data, theme)

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

        _add_decorative_accents(slide, layout_type, has_background=bool(background_stream), theme=theme)
        _draw_title(slide, slide_data.get("title", f"Slide {slide_index}"), bool(background_stream), theme)

        image_asset = _resolve_foreground_asset(slide_data, harvested_images)
        blocks = _normalize_blocks(slide_data)
        layout_type = _normalize_layout_type(layout_type, image_asset, blocks, slide_data)

        _render_layout(slide, layout_type, blocks, image_asset, bool(background_stream), theme, slide_data)

    _add_thank_you_slide(prs)
    prs.save(output_filename)
    return output_filename


def _normalize_key(value):
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


def _resolve_deck_title(master_blueprint_json, output_filename):
    if master_blueprint_json and isinstance(master_blueprint_json[0], dict):
        first_title = str(master_blueprint_json[0].get("title") or "").strip()
        if first_title:
            return first_title
    stem = os.path.splitext(os.path.basename(output_filename))[0]
    stem = stem.replace("_", " ").replace("-", " ").strip()
    return stem or "Presentation"


def _add_cover_slide(prs, deck_title):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme = THEMES["midnight"]

    bg = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(0),
        Inches(0),
        Inches(SLIDE_WIDTH),
        Inches(SLIDE_HEIGHT),
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(245, 247, 255)
    bg.line.fill.background()

    _render_background_design(
        slide,
        [
            {"type": "circle", "x": -1.0, "y": -1.1, "w": 4.8, "h": 4.8, "fill": "#C9D2FF", "transparency": 0.18},
            {"type": "rounded_rect", "x": 0.95, "y": 1.3, "w": 11.1, "h": 4.7, "fill": "#E7EBFF", "transparency": 0.22},
            {"type": "arc_band", "x": 9.4, "y": 4.7, "w": 4.2, "h": 3.0, "fill": "#6D78FF", "transparency": 0.18},
        ],
        theme,
    )

    accent = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(0.95),
        Inches(1.12),
        Inches(1.6),
        Inches(0.14),
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = theme["accent"]
    accent.line.fill.background()

    title_box = slide.shapes.add_textbox(Inches(0.98), Inches(1.55), Inches(8.6), Inches(2.2))
    tf = title_box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    p = tf.paragraphs[0]
    p.text = deck_title
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = PALETTE["navy"]

    subtitle_box = slide.shapes.add_textbox(Inches(1.02), Inches(3.85), Inches(6.7), Inches(0.9))
    stf = subtitle_box.text_frame
    stf.word_wrap = True
    stf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    sp = stf.paragraphs[0]
    sp.text = "Strategic presentation generated with structured storytelling and editorial slide design"
    sp.font.size = Pt(16)
    sp.font.color.rgb = PALETTE["slate"]

    _add_decorative_visual_panel(
        slide,
        theme,
        left=Inches(8.7),
        top=Inches(1.65),
        width=Inches(3.1),
        height=Inches(2.6),
    )

    return slide


def _add_thank_you_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme = THEMES["midnight"]

    bg = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(0),
        Inches(0),
        Inches(SLIDE_WIDTH),
        Inches(SLIDE_HEIGHT),
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(244, 246, 255)
    bg.line.fill.background()

    _render_background_design(
        slide,
        [
            {"type": "circle", "x": -0.9, "y": -0.9, "w": 4.0, "h": 4.0, "fill": "#CDD6FF", "transparency": 0.18},
            {"type": "rounded_rect", "x": 1.0, "y": 1.4, "w": 11.0, "h": 4.7, "fill": "#E9ECFF", "transparency": 0.2},
            {"type": "arc_band", "x": 9.5, "y": 4.8, "w": 4.0, "h": 2.8, "fill": "#6F79FF", "transparency": 0.18},
        ],
        theme,
    )

    accent = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(5.22),
        Inches(2.0),
        Inches(2.9),
        Inches(0.15),
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = theme["accent"]
    accent.line.fill.background()

    title_box = slide.shapes.add_textbox(Inches(2.0), Inches(2.45), Inches(9.4), Inches(1.8))
    tf = title_box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    p = tf.paragraphs[0]
    p.text = "Thank You"
    p.font.size = Pt(30)
    p.font.bold = True
    p.font.color.rgb = PALETTE["navy"]
    p.alignment = PP_ALIGN.CENTER

    _add_decorative_visual_panel(
        slide,
        theme,
        left=Inches(4.95),
        top=Inches(4.55),
        width=Inches(3.1),
        height=Inches(1.35),
    )

    return slide


def _resolve_theme(slide_data, slide_index):
    explicit_theme = str(slide_data.get("theme_variant") or "").strip().lower()
    if explicit_theme in THEMES:
        return THEMES[explicit_theme]

    theme_names = sorted(THEMES.keys())
    seed_basis = f"{slide_index}:{slide_data.get('title', '')}:{slide_data.get('layout_type', '')}"
    seed_value = int(hashlib.sha256(seed_basis.encode("utf-8")).hexdigest()[:8], 16)
    theme_name = theme_names[seed_value % len(theme_names)]
    return THEMES[theme_name]


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


def _resolve_foreground_asset(slide_data, harvested_images):
    existing_chart_info = slide_data.get("existing_chart", {}) or {}
    topic_ref = existing_chart_info.get("chart_topic_reference")
    best_match_key = find_best_match(topic_ref, list(harvested_images.keys()))
    decoration_stream = None

    if existing_chart_info.get("has_existing_chart") and best_match_key:
        stream = _first_image_stream(harvested_images.get(best_match_key))
        foreground_info = slide_data.get("foreground_image", {}) or {}
        prompt = foreground_info.get("foreground_image_prompt")
        if foreground_info.get("has_foreground_image") and prompt:
            decoration_stream = _generate_image_from_prompt(
                f"{prompt}, decorative supporting visual, small elegant accent asset",
                width=768,
                height=768,
                timeout=45,
            )
        if stream:
            print(f"Using extracted chart for topic: {best_match_key}")
            return {
                "stream": stream,
                "kind": "chart",
                "source": "extracted",
                "topic": best_match_key,
                "decoration_stream": decoration_stream,
            }

    foreground_info = slide_data.get("foreground_image", {}) or {}
    if foreground_info.get("has_foreground_image"):
        prompt = foreground_info.get("foreground_image_prompt")
        if prompt:
            stream = _generate_image_from_prompt(
                f"{prompt}, isolated subject, transparent style, presentation-ready asset",
                width=1024,
                height=768,
                timeout=45,
            )
            if stream:
                return {
                    "stream": stream,
                    "kind": "illustration",
                    "source": "generated",
                    "topic": topic_ref,
                    "decoration_stream": None,
                }

    visual_priority = str(slide_data.get("visual_priority") or "").strip().lower()
    if visual_priority in {"image_focus", "balanced"} and not existing_chart_info.get("has_existing_chart"):
        generated_prompt = _build_visual_prompt_from_slide(slide_data)
        stream = _generate_image_from_prompt(
            generated_prompt,
            width=1024,
            height=768,
            timeout=45,
        )
        if stream:
            return {
                "stream": stream,
                "kind": "illustration",
                "source": "auto-generated",
                "topic": topic_ref,
                "decoration_stream": None,
            }

    fallback_candidates = []
    if topic_ref and best_match_key:
        fallback_candidates.append(harvested_images.get(best_match_key))
    for candidate in fallback_candidates:
        stream = _first_image_stream(candidate)
        if stream:
            return {
                "stream": stream,
                "kind": "chart",
                "source": "fallback",
                "topic": topic_ref,
                "decoration_stream": None,
            }

    return {"stream": None, "kind": "none", "source": "none", "topic": topic_ref, "decoration_stream": None}


def _generate_image_from_prompt(prompt, width, height, timeout):
    styled_prompt = f"{prompt}, {IMAGE_STYLE_SUFFIX}"
    cache_key = (styled_prompt, width, height)
    if cache_key in IMAGE_CACHE:
        return io.BytesIO(IMAGE_CACHE[cache_key])

    fallback_stream = _build_prompt_fallback_image(styled_prompt, width, height)
    if fallback_stream:
        fallback_bytes = fallback_stream.getvalue()
        IMAGE_CACHE[cache_key] = fallback_bytes
        return io.BytesIO(fallback_bytes)
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

    summary = str(slide_data.get("summary") or "").strip()
    bullets = slide_data.get("bullets", []) or []
    if isinstance(bullets, list) and (summary or bullets):
        normalized = [
            {
                "block_title": "Key Takeaway",
                "bullets": [str(item).strip() for item in bullets[:4] if str(item).strip()],
                "body": summary,
            }
        ]
        metrics = slide_data.get("highlight_metrics", []) or []
        metrics = [str(item).strip() for item in metrics[:4] if str(item).strip()]
        if metrics:
            normalized.append(
                {
                    "block_title": "Key Metrics",
                    "bullets": metrics,
                    "body": "",
                }
            )
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


def _normalize_layout_type(layout_type, image_asset, blocks, slide_data):
    visual_priority = str(slide_data.get("visual_priority") or "").strip().lower()
    recommended_visual = str(slide_data.get("recommended_visual") or "").strip().lower()
    title = str(slide_data.get("title") or "").lower()

    if layout_type not in {
        "top_image_bottom_cards",
        "hero_background_with_floating_cards",
        "split_left_image_right_cards",
        "diagonal_text_split",
        "multi_column_cards",
        "chart_focus_cards",
        "metric_dashboard",
        "timeline_roadmap",
    }:
        layout_type = "split_left_image_right_cards"

    if image_asset.get("kind") == "chart":
        if layout_type in {"multi_column_cards", "hero_background_with_floating_cards"}:
            return "chart_focus_cards"
        if visual_priority == "chart_focus":
            return "chart_focus_cards"
    if recommended_visual == "dashboard" or len(slide_data.get("highlight_metrics", []) or []) >= 3:
        return "metric_dashboard"
    if recommended_visual == "timeline" or "roadmap" in title or "timeline" in title:
        return "timeline_roadmap"
    return layout_type


def _render_layout(slide, layout_type, blocks, image_asset, has_background, theme, slide_data):
    if layout_type == "metric_dashboard":
        _layout_metric_dashboard(slide, blocks, image_asset, theme, slide_data)
    elif layout_type == "timeline_roadmap":
        _layout_timeline_roadmap(slide, blocks, image_asset, theme)
    elif layout_type == "top_image_bottom_cards":
        _layout_top_image_bottom_cards(slide, blocks, image_asset, theme)
    elif layout_type == "hero_background_with_floating_cards":
        _layout_hero_background_with_floating_cards(slide, blocks, image_asset, has_background, theme)
    elif layout_type == "multi_column_cards":
        _layout_multi_column_cards(slide, blocks, image_asset, has_background, theme)
    elif layout_type == "chart_focus_cards":
        _layout_chart_focus_cards(slide, blocks, image_asset, theme)
    elif layout_type == "diagonal_text_split":
        _layout_diagonal_text_split(slide, blocks, image_asset, theme)
    else:
        _layout_split_left_image_right_cards(slide, blocks, image_asset, theme)


def _layout_metric_dashboard(slide, blocks, image_asset, theme, slide_data):
    metrics = [str(item).strip() for item in (slide_data.get("highlight_metrics") or []) if str(item).strip()][:4]
    if metrics:
        metric_specs = _horizontal_card_specs(len(metrics), Inches(1.45), Inches(1.15), gutter=0.18)
        for index, metric in enumerate(metrics):
            left, top, width, height = metric_specs[index]
            _add_metric_card(slide, left, top, width, height, metric, theme)

    image_stream = image_asset.get("stream")
    if image_stream:
        _add_picture_contain(slide, image_stream, Inches(0.78), Inches(3.0), Inches(5.35), Inches(2.85))
    else:
        _add_decorative_visual_panel(slide, theme, left=Inches(0.95), top=Inches(3.2), width=Inches(4.9), height=Inches(2.45))

    content_specs = [
        (Inches(6.45), Inches(2.15), Inches(6.0), Inches(1.65)),
        (Inches(6.45), Inches(4.0), Inches(2.82), Inches(2.0)),
        (Inches(9.63), Inches(4.0), Inches(2.82), Inches(2.0)),
    ]
    for index, block in enumerate(blocks[: len(content_specs)]):
        left, top, width, height = content_specs[index]
        _add_card(slide, left, top, width, height, block, dark=False, theme=theme)


def _layout_timeline_roadmap(slide, blocks, image_asset, theme):
    line = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(1.0), Inches(3.35), Inches(11.0), Inches(0.08))
    line.fill.solid()
    line.fill.fore_color.rgb = theme["accent"]
    line.line.fill.background()

    milestones = blocks[:4] if blocks else [{"block_title": "Milestone", "body": "", "bullets": []}]
    count = len(milestones)
    for index, block in enumerate(milestones):
        center_x = 1.35 + (index * (10.0 / max(count - 1, 1)))
        node = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(center_x), Inches(3.08), Inches(0.42), Inches(0.42))
        node.fill.solid()
        node.fill.fore_color.rgb = theme["accent_dark"]
        node.line.fill.background()

        top = Inches(1.7 if index % 2 == 0 else 3.75)
        _add_card(slide, Inches(center_x - 0.7), top, Inches(2.0), Inches(1.4), block, dark=False, theme=theme)

    image_stream = image_asset.get("stream")
    if image_stream and image_asset.get("kind") != "chart":
        _add_picture_contain(slide, image_stream, Inches(10.4), Inches(0.95), Inches(1.9), Inches(1.5))


def _layout_top_image_bottom_cards(slide, blocks, image_asset, theme):
    image_stream = image_asset.get("stream")
    if image_stream:
        _add_picture_contain(slide, image_stream, Inches(0.75), Inches(1.3), Inches(11.85), Inches(2.6))
    else:
        _add_placeholder_panel(slide, Inches(0.95), Inches(1.45), Inches(11.45), Inches(2.25), "Visual summary")

    for index, (left, top, width, height) in enumerate(_horizontal_card_specs(len(blocks), Inches(4.25), Inches(2.2))):
        _add_card(slide, left, top, width, height, blocks[index], dark=False, theme=theme)


def _layout_hero_background_with_floating_cards(slide, blocks, image_asset, has_background, theme):
    image_stream = image_asset.get("stream")
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
        _add_card(slide, left, top, width, height, block, dark=True, theme=theme)


def _layout_multi_column_cards(slide, blocks, image_asset, has_background, theme):
    image_stream = image_asset.get("stream")
    density = _estimate_content_density(blocks)
    if image_stream and image_asset.get("kind") != "chart":
        _layout_multi_column_with_visual(slide, blocks, image_asset, has_background, theme, density)
        return

    if not has_background:
        _add_band(slide, Inches(0.6), Inches(1.35), Inches(12.1), Inches(5.05), theme["band"], transparency=0.28)

    if not image_stream and density == "light":
        _add_decorative_visual_panel(slide, theme)

    card_top = Inches(2.0 if density == "light" else 1.75)
    card_height = Inches(3.25 if density == "light" else 4.1)
    for index, (left, top, width, height) in enumerate(_horizontal_card_specs(len(blocks), card_top, card_height, gutter=0.22)):
        _add_card(slide, left, top, width, height, blocks[index], dark=False, theme=theme)


def _layout_multi_column_with_visual(slide, blocks, image_asset, has_background, theme, density):
    image_stream = image_asset.get("stream")

    if not has_background:
        _add_band(slide, Inches(0.55), Inches(1.28), Inches(12.2), Inches(5.18), theme["band"], transparency=0.18)

    image_frame_left = Inches(7.85)
    image_frame_top = Inches(1.48)
    image_frame_width = Inches(4.35)
    image_frame_height = Inches(4.25 if density != "light" else 3.95)

    _add_visual_frame(slide, image_frame_left, image_frame_top, image_frame_width, image_frame_height, theme)
    _add_picture_cover(slide, image_stream, image_frame_left, image_frame_top, image_frame_width, image_frame_height)

    max_cards = min(len(blocks), 3)
    card_left = Inches(0.72)
    card_width = Inches(6.55)
    if max_cards == 1:
        specs = [(card_left, Inches(1.75), card_width, Inches(2.15))]
    elif max_cards == 2:
        specs = [
            (card_left, Inches(1.62), card_width, Inches(2.18)),
            (card_left + Inches(0.32), Inches(4.02), card_width - Inches(0.32), Inches(1.95)),
        ]
    else:
        specs = [
            (card_left, Inches(1.55), card_width, Inches(1.7)),
            (card_left + Inches(0.2), Inches(3.38), card_width - Inches(0.2), Inches(1.55)),
            (card_left + Inches(0.4), Inches(5.03), card_width - Inches(0.4), Inches(1.25)),
        ]

    for index in range(max_cards):
        left, top, width, height = specs[index]
        _add_card(slide, left, top, width, height, blocks[index], dark=False, theme=theme)

    if len(blocks) == 0:
        _add_placeholder_panel(slide, Inches(0.85), Inches(2.0), Inches(6.1), Inches(2.4), "Key insights")


def _layout_chart_focus_cards(slide, blocks, image_asset, theme):
    image_stream = image_asset.get("stream")
    decoration_stream = image_asset.get("decoration_stream")
    chart_left = Inches(0.75)
    chart_top = Inches(1.45)
    chart_width = Inches(5.9)
    chart_height = Inches(4.6)

    if image_stream:
        _add_picture_contain(slide, image_stream, chart_left, chart_top, chart_width, chart_height)
    else:
        _add_placeholder_panel(slide, chart_left, chart_top, chart_width, chart_height, "Primary chart")

    density = _estimate_content_density(blocks)
    card_height = Inches(1.55 if density == "light" else 1.95)
    start_top = Inches(1.55)
    for index, block in enumerate(blocks[:3]):
        _add_card(
            slide,
            Inches(7.0),
            Inches(1.55 + (index * (card_height.inches + 0.22))),
            Inches(5.55),
            card_height,
            block,
            dark=False,
            theme=theme,
        )

    if decoration_stream:
        _add_picture_contain(slide, decoration_stream, Inches(10.15), Inches(0.95), Inches(2.05), Inches(1.25))

    if density == "light" and image_asset.get("kind") == "chart":
        _add_data_caption(slide, "Chart-led slide layout selected automatically", theme)


def _layout_diagonal_text_split(slide, blocks, image_asset, theme):
    image_stream = image_asset.get("stream")
    if blocks:
        _add_card(slide, Inches(0.7), Inches(1.45), Inches(4.15), Inches(2.05), blocks[0], dark=False, theme=theme)
    if len(blocks) > 1:
        _add_card(slide, Inches(8.35), Inches(4.0), Inches(4.2), Inches(2.05), blocks[1], dark=False, theme=theme)
    if len(blocks) > 2:
        _add_card(slide, Inches(0.95), Inches(4.25), Inches(3.5), Inches(1.45), blocks[2], dark=True, theme=theme)

    if image_stream:
        _add_picture_contain(slide, image_stream, Inches(4.55), Inches(2.0), Inches(4.2), Inches(3.05))
    else:
        _add_placeholder_panel(slide, Inches(4.85), Inches(2.35), Inches(3.65), Inches(2.35), "Central visual")


def _layout_split_left_image_right_cards(slide, blocks, image_asset, theme):
    image_stream = image_asset.get("stream")
    if image_stream:
        _add_picture_contain(slide, image_stream, Inches(0.7), Inches(1.45), Inches(6.05), Inches(4.9))
    else:
        _add_placeholder_panel(slide, Inches(0.9), Inches(1.75), Inches(5.6), Inches(4.3), "Primary visual")

    density = _estimate_content_density(blocks)
    max_cards = min(len(blocks), 3)
    card_height = (4.2 if density == "light" else 4.9) / max(max_cards, 1)
    for index in range(max_cards):
        _add_card(
            slide,
            Inches(7.1),
            Inches(1.45 + (index * card_height)),
            Inches(5.45),
            Inches(card_height - 0.15),
            blocks[index],
            dark=False,
            theme=theme,
        )

    if density == "light" and image_asset.get("kind") == "illustration":
        _add_data_caption(slide, "Concept visual paired with concise insights", theme)


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


def _estimate_content_density(blocks):
    total_items = 0
    total_chars = 0
    for block in blocks[:4]:
        total_items += len(block.get("bullets", []) or [])
        total_chars += len(block.get("body", "") or "")
        total_chars += len(block.get("block_title", "") or "")
        total_chars += sum(len(str(item)) for item in block.get("bullets", []) or [])

    if total_items <= 4 and total_chars < 220:
        return "light"
    if total_items >= 9 or total_chars > 520:
        return "heavy"
    return "medium"


def _paint_base_background(slide, slide_data, theme):
    design = slide_data.get("background_design", {}) or {}
    base_fill = _coerce_rgb(design.get("base_fill")) or theme["base"]
    base = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(0),
        Inches(0),
        Inches(SLIDE_WIDTH),
        Inches(SLIDE_HEIGHT),
    )
    base.fill.solid()
    base.fill.fore_color.rgb = base_fill
    base.line.fill.background()
    elements = design.get("elements")
    if isinstance(elements, list) and elements:
        _render_background_design(slide, elements, theme)
    else:
        glow = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.OVAL,
            Inches(-0.9),
            Inches(-1.2),
            Inches(4.2),
            Inches(4.2),
        )
        glow.fill.solid()
        glow.fill.fore_color.rgb = theme["glow"]
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
        corner.fill.fore_color.rgb = theme["corner"]
        corner.fill.transparency = 0.58
        corner.line.fill.background()


def _render_background_design(slide, elements, theme):
    for element in elements[:6]:
        if not isinstance(element, dict):
            continue

        shape_type = str(element.get("type") or "").strip().lower()
        x = Inches(float(element.get("x", 0)))
        y = Inches(float(element.get("y", 0)))
        w = Inches(max(0.1, float(element.get("w", 1.0))))
        h = Inches(max(0.1, float(element.get("h", 1.0))))
        fill_color = _coerce_rgb(element.get("fill")) or theme["glow"]
        transparency = max(0.0, min(0.95, float(element.get("transparency", 0.2))))

        if shape_type == "circle":
            shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, x, y, w, h)
        elif shape_type == "rounded_rect":
            shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, y, w, h)
        elif shape_type == "arc_band":
            shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ARC, x, y, w, h)
            line = shape.line
            line.color.rgb = fill_color
            line.width = Pt(10)
            shape.fill.background()
            continue
        else:
            shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, x, y, w, h)

        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
        shape.fill.transparency = transparency
        shape.line.fill.background()


def _coerce_rgb(value):
    if isinstance(value, RGBColor):
        return value
    if not value:
        return None

    text = str(value).strip()
    if text.startswith("#"):
        text = text[1:]
    if len(text) != 6:
        return None
    try:
        return RGBColor(int(text[0:2], 16), int(text[2:4], 16), int(text[4:6], 16))
    except ValueError:
        return None


def _build_visual_prompt_from_slide(slide_data):
    title = str(slide_data.get("title") or "Business strategy concept").strip()
    blocks = slide_data.get("content_blocks") or []
    keywords = []
    for block in blocks[:2]:
        if not isinstance(block, dict):
            continue
        block_title = str(block.get("block_title") or "").strip()
        if block_title:
            keywords.append(block_title)
        for bullet in block.get("bullets", [])[:2]:
            bullet_text = str(bullet).strip()
            if bullet_text:
                keywords.append(bullet_text)
    keyword_text = ", ".join(keywords[:4])
    if keyword_text:
        return f"{title}, {keyword_text}, abstract strategic illustration, editorial business storytelling"
    return f"{title}, abstract strategic illustration, editorial business storytelling"


def _add_decorative_accents(slide, layout_type, has_background, theme):
    accent = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(0.55),
        Inches(0.62),
        Inches(1.15),
        Inches(0.1),
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = theme["accent"]
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


def _draw_title(slide, title_text, on_busy_background, theme):
    title_box = slide.shapes.add_textbox(Inches(0.78), Inches(0.62), Inches(11.7), Inches(0.95))
    title_box.text_frame.word_wrap = True
    title_box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = title_box.text_frame.paragraphs[0]
    p.text = str(title_text or "Untitled Slide")
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = PALETTE["navy"] if not on_busy_background else PALETTE["white"]
    p.alignment = PP_ALIGN.LEFT


def _add_card(slide, left, top, width, height, block_data, dark=False, theme=None):
    theme = theme or THEMES["ocean"]
    card = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = theme["accent_dark"] if dark else theme["card_fill"]
    card.fill.transparency = 0.1 if dark else 0.02
    card.line.color.rgb = theme["glow"] if dark else PALETTE["line"]
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
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.margin_left = Inches(0.22)
    tf.margin_right = Inches(0.22)
    tf.margin_top = Inches(0.16)
    tf.margin_bottom = Inches(0.14)
    tf.vertical_anchor = MSO_ANCHOR.TOP

    content_length = len(str(block_data.get("body") or ""))
    content_length += sum(len(str(item)) for item in (block_data.get("bullets", []) or []))
    content_length += len(str(block_data.get("block_title") or ""))
    compact = content_length > 260

    title = str(block_data.get("block_title") or "Insight").strip()
    title_para = tf.paragraphs[0]
    title_para.text = title
    title_para.font.size = Pt(16 if compact else (18 if dark else 17))
    title_para.font.bold = True
    title_para.font.color.rgb = PALETTE["white"] if dark else PALETTE["navy"]
    title_para.space_after = Pt(6 if compact else 8)

    body = str(block_data.get("body") or "").strip()
    if body:
        body_para = tf.add_paragraph()
        body_para.text = body
        body_para.font.size = Pt(10 if compact else 11.5)
        body_para.font.color.rgb = RGBColor(226, 232, 240) if dark else PALETTE["slate"]
        body_para.space_after = Pt(5 if compact else 8)

    bullets = block_data.get("bullets", []) or []
    bullet_limit = 6 if compact else 5
    for bullet in bullets[:bullet_limit]:
        bullet_para = tf.add_paragraph()
        bullet_para.text = f"- {str(bullet)}"
        bullet_para.level = 0
        bullet_para.font.size = Pt(9.5 if compact else 11.5)
        bullet_para.font.color.rgb = RGBColor(226, 232, 240) if dark else PALETTE["ink"]
        bullet_para.space_after = Pt(2 if compact else 4)


def _add_metric_card(slide, left, top, width, height, metric, theme):
    card = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = theme["accent_dark"]
    card.fill.transparency = 0.02
    card.line.fill.background()
    tf = card.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.18)
    tf.margin_right = Inches(0.18)
    tf.margin_top = Inches(0.12)
    p = tf.paragraphs[0]
    p.text = metric
    p.font.bold = True
    p.font.size = Pt(16)
    p.font.color.rgb = PALETTE["white"]
    p.alignment = PP_ALIGN.CENTER
    return card


def _add_visual_frame(slide, left, top, width, height, theme):
    shadow = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        left + Inches(0.08),
        top + Inches(0.1),
        width,
        height,
    )
    shadow.fill.solid()
    shadow.fill.fore_color.rgb = RGBColor(148, 163, 184)
    shadow.fill.transparency = 0.72
    shadow.line.fill.background()

    frame = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, left, top, width, height)
    frame.fill.solid()
    frame.fill.fore_color.rgb = PALETTE["white"]
    frame.line.color.rgb = theme["glow"]
    frame.line.width = Pt(1.2)
    return frame


def _add_decorative_visual_panel(slide, theme, left=Inches(9.1), top=Inches(1.55), width=Inches(2.75), height=Inches(1.95)):
    panel = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        left,
        top,
        width,
        height,
    )
    panel.fill.solid()
    panel.fill.fore_color.rgb = theme["glow"]
    panel.fill.transparency = 0.18
    panel.line.fill.background()

    orbit = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.OVAL,
        left + Inches(0.45),
        top + Inches(0.3),
        Inches(min(1.15, width.inches * 0.42)),
        Inches(min(1.15, height.inches * 0.58)),
    )
    orbit.fill.solid()
    orbit.fill.fore_color.rgb = theme["accent"]
    orbit.fill.transparency = 0.08
    orbit.line.fill.background()

    orbit2 = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.OVAL,
        left + Inches(min(1.35, width.inches * 0.55)),
        top + Inches(min(0.55, height.inches * 0.28)),
        Inches(min(0.78, width.inches * 0.28)),
        Inches(min(0.78, height.inches * 0.4)),
    )
    orbit2.fill.solid()
    orbit2.fill.fore_color.rgb = theme["accent_dark"]
    orbit2.fill.transparency = 0.15
    orbit2.line.fill.background()


def _add_data_caption(slide, text, theme):
    caption = slide.shapes.add_textbox(Inches(0.9), Inches(6.35), Inches(4.2), Inches(0.35))
    tf = caption.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(9.5)
    p.font.color.rgb = theme["accent_dark"]
    p.alignment = PP_ALIGN.LEFT


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


def _build_prompt_fallback_image(prompt, width, height):
    if Image is None or ImageDraw is None:
        return None

    image_width = max(800, int(width))
    image_height = max(450, int(height))
    seed = int(hashlib.sha256(prompt.encode("utf-8")).hexdigest()[:16], 16)
    rng = random.Random(seed)

    canvas = Image.new("RGB", (image_width, image_height), (14, 23, 42))
    draw = ImageDraw.Draw(canvas, "RGBA")

    gradient_colors = [
        (14, 116, 144),
        (30, 64, 175),
        (15, 23, 42),
        (56, 189, 248),
    ]
    for y in range(image_height):
        t = y / max(image_height - 1, 1)
        start = gradient_colors[0]
        end = gradient_colors[1] if t < 0.5 else gradient_colors[2]
        mix = t * 2 if t < 0.5 else (t - 0.5) * 2
        r = int(start[0] + (end[0] - start[0]) * mix)
        g = int(start[1] + (end[1] - start[1]) * mix)
        b = int(start[2] + (end[2] - start[2]) * mix)
        draw.line((0, y, image_width, y), fill=(r, g, b, 255))

    for _ in range(12):
        x0 = rng.randint(-image_width // 8, image_width)
        y0 = rng.randint(-image_height // 8, image_height)
        w = rng.randint(image_width // 8, image_width // 2)
        h = rng.randint(image_height // 8, image_height // 2)
        color = rng.choice(
            [
                (125, 211, 252, 55),
                (56, 189, 248, 70),
                (191, 219, 254, 45),
                (45, 212, 191, 50),
            ]
        )
        draw.rounded_rectangle((x0, y0, x0 + w, y0 + h), radius=rng.randint(18, 60), fill=color)

    for _ in range(9):
        x1 = rng.randint(0, image_width)
        y1 = rng.randint(0, image_height)
        x2 = rng.randint(0, image_width)
        y2 = rng.randint(0, image_height)
        draw.line((x1, y1, x2, y2), fill=(226, 232, 240, 90), width=rng.randint(3, 8))

    _draw_prompt_iconography(draw, prompt, image_width, image_height, rng)

    label = _prompt_to_visual_label(prompt)
    font = None
    if ImageFont is not None:
        try:
            font = ImageFont.truetype("arial.ttf", max(28, image_width // 30))
        except Exception:
            font = ImageFont.load_default()

    if font:
        bbox = draw.textbbox((0, 0), label, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        text_x = (image_width - text_width) // 2
        text_y = max(24, image_height - text_height - 36)
        draw.rounded_rectangle(
            (text_x - 20, text_y - 12, text_x + text_width + 20, text_y + text_height + 12),
            radius=18,
            fill=(15, 23, 42, 180),
        )
        draw.text((text_x, text_y), label, fill=(255, 255, 255), font=font)

    stream = io.BytesIO()
    canvas.save(stream, format="PNG")
    stream.seek(0)
    return stream


def _draw_prompt_iconography(draw, prompt, image_width, image_height, rng):
    prompt_lower = str(prompt or "").lower()
    center_x = image_width // 2
    center_y = image_height // 2 - 30

    if any(word in prompt_lower for word in ["roadmap", "timeline", "phase", "implementation", "milestone"]):
        draw.line((120, center_y, image_width - 120, center_y), fill=(255, 255, 255, 170), width=8)
        for idx in range(4):
            x = 180 + idx * ((image_width - 360) // 3)
            draw.ellipse((x - 22, center_y - 22, x + 22, center_y + 22), fill=(56, 189, 248, 220))
            draw.rounded_rectangle((x - 60, center_y - 120 + (idx % 2) * 130, x + 60, center_y - 55 + (idx % 2) * 130), radius=16, fill=(15, 23, 42, 180))
    elif any(word in prompt_lower for word in ["dashboard", "metric", "revenue", "market", "finance", "kpi"]):
        base_y = center_y + 110
        bar_width = 70
        gaps = 45
        heights = [110, 165, 235]
        start_x = center_x - ((3 * bar_width + 2 * gaps) // 2)
        for idx, bar_height in enumerate(heights):
            x0 = start_x + idx * (bar_width + gaps)
            draw.rounded_rectangle((x0, base_y - bar_height, x0 + bar_width, base_y), radius=18, fill=(56, 189, 248, 220))
        draw.line((center_x - 170, base_y - 20, center_x + 170, base_y - 190), fill=(255, 255, 255, 185), width=9)
    elif any(word in prompt_lower for word in ["network", "platform", "ecosystem", "ai", "data", "cloud"]):
        nodes = []
        for _ in range(6):
            x = rng.randint(center_x - 180, center_x + 180)
            y = rng.randint(center_y - 120, center_y + 120)
            nodes.append((x, y))
        for i, (x1, y1) in enumerate(nodes):
            for x2, y2 in nodes[i + 1 :]:
                if abs(x1 - x2) + abs(y1 - y2) < 280:
                    draw.line((x1, y1, x2, y2), fill=(191, 219, 254, 130), width=5)
        for x, y in nodes:
            draw.ellipse((x - 24, y - 24, x + 24, y + 24), fill=(45, 212, 191, 220))
    else:
        draw.rounded_rectangle((center_x - 180, center_y - 120, center_x + 180, center_y + 120), radius=28, fill=(15, 23, 42, 140))
        for idx in range(3):
            y = center_y - 45 + idx * 48
            draw.rounded_rectangle((center_x - 130, y, center_x + 130, y + 18), radius=9, fill=(255, 255, 255, 170))


def _prompt_to_visual_label(prompt):
    cleaned = re.sub(r"\s+", " ", str(prompt or "")).strip()
    fragments = [fragment.strip(" ,.-") for fragment in cleaned.split(",") if fragment.strip()]
    for fragment in fragments:
        if len(fragment) >= 12:
            words = fragment.split()
            return " ".join(words[:6]).title()
    return "Concept Visual"
import json
import os
from dotenv import load_dotenv
from google import genai
from google.genai import types

load_dotenv()

def _get_client() -> genai.Client:
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY is not set.")
    return genai.Client(api_key=api_key)

def design_slide_layouts(summarized_json_string: str, available_topics: list, progress_callback=None) -> str:
    """Agent 2: The Art Director. Assigns ultra-premium magazine layouts."""
    client = _get_client()
    
    prompt = f"""
    You are an elite presentation art director designing a visually stunning, premium consulting-style pitch deck.
    I will give you a JSON array of raw slide data. Your job is to transform it into a polished visual blueprint for a PowerPoint compiler.

    DESIGN GOAL:
    Make the deck feel premium, modern, bold, and cohesive.
    Use clean editorial composition, asymmetry, strong whitespace, clear hierarchy, and visual variation across slides.
    The style should feel like a high-end strategy deck mixed with stylized product storytelling.

    AVAILABLE LAYOUT TYPES:
    1. "top_image_bottom_cards" = strong hero visual on top, 2 to 3 content cards below
    2. "hero_background_with_floating_cards" = full visual atmosphere in background with floating cards
    3. "split_left_image_right_cards" = image or chart on left, stacked cards on right
    4. "diagonal_text_split" = dynamic diagonal composition with center visual
    5. "multi_column_cards" = elegant grid of 2 to 4 vertical cards, background can be subtle
    6. "metric_dashboard" = top metric band with bottom visual or insight cards
    7. "timeline_roadmap" = horizontal roadmap or phased implementation story with milestone cards

    THEME VARIANTS:
    Choose one per slide from:
    - "ocean"
    - "midnight"
    - "emerald"
    - "sunset"

    VISUAL PRIORITY:
    For each slide also choose one:
    - "chart_focus" when a real extracted chart should dominate the slide
    - "image_focus" when an abstract illustration should dominate
    - "balanced" when text and visuals should share space
    - "text_focus" when strong cards matter more than imagery

    LAYOUT SELECTION RULE:
    If recommended_visual is:
    - "dashboard" or the slide has many highlight_metrics, strongly prefer "metric_dashboard"
    - "timeline" or the slide describes phases/years/roadmap, strongly prefer "timeline_roadmap"
    - "framework", "comparison", or "process", prefer layouts that keep multiple structured cards visible at once

    IMAGE AND CHART RULES:
    I already have real, extracted charts for these topics:
    {available_topics}

    If a slide topic closely matches one of those extracted chart topics, set:
    "existing_chart": {{
        "has_existing_chart": true,
        "chart_topic_reference": "Exact Topic Name"
    }}
    and prefer the extracted chart over AI-generated imagery.
    When you use an extracted chart, choose "visual_priority": "chart_focus" unless there is a clear reason not to.

    If the incoming raw slide data already contains an "existing_chart" object, you MUST preserve it.

    IMPORTANT RATE-LIMIT RULE:
    Do not request too many generated images.
    Prefer at most ONE generated image per slide.
    Many slides should use either:
    - only a background image, OR
    - only a foreground image, OR
    - no generated image at all if the layout is already strong with cards and a chart.

    If a slide uses an extracted chart, usually set "has_foreground_image": false unless a small conceptual visual is truly necessary.
    Use background images sparingly. Reserve them for title slides, transition slides, visionary slides, or slides that benefit from atmosphere.

    IMAGE PROMPT STYLE:
    Prompts must be short, vivid, and visual.
    Use stylized abstract 3D, geometric illustration, layered shapes, isometric forms, glowing paths, symbolic systems, or elegant conceptual scenes.
    Avoid photorealistic humans, realistic photography, crowded scenes, tiny details, or text inside the image.
    Keep the entire deck in one cohesive visual family: abstract, premium, slightly futuristic, corporate-editorial, not cartoonish.

    BACKGROUND DESIGN RULES:
    For every slide, create a "background_design" object that describes the decorative background the compiler should render.
    Use it to create variety with arcs, circles, soft bands, rounded rectangles, outlines, light panels, diagonal bars, or framing elements.
    Do not repeat the same composition on every slide.
    Vary placement, scale, and density based on the slide's layout and content.
    Keep it elegant and subtle enough that content remains readable.
    Use 2 to 5 elements.

    Each element must use this schema:
    {{
        "type": "circle" | "rounded_rect" | "rect" | "arc_band",
        "x": 0.0,
        "y": 0.0,
        "w": 3.0,
        "h": 2.0,
        "fill": "#DCEEFF",
        "transparency": 0.25
    }}

    GENERATED IMAGE RULE:
    If visual_priority is "image_focus" and the slide is not chart-led, you should usually set "foreground_image.has_foreground_image" to true.
    If the slide is a hero or transition-style slide, you should usually set either a background image or a foreground image, not neither.

    CONTENT BLOCK RULES:
    Group the text into 2 to 5 meaningful content blocks.
    Each block should have a short block_title and concise bullets.
    If the incoming raw slide data contains "summary" or "highlight_metrics", use them.
    Turn the summary into a short body/supporting sentence where helpful.
    If metrics are present, dedicate one content block or card to them so key numbers become visually dominant.
    Prioritize readability and visual balance, not dumping all text into one area.

    CHART + DECORATION RULE:
    Even when a slide has an extracted chart, you should often still add visual atmosphere.
    In chart-led slides, prefer one of these:
    - a subtle generated background image, or
    - a small decorative foreground image that does not compete with the chart
    Do not leave chart slides feeling dry or text-only.
    
    Return ONLY a valid JSON array of objects. 
    Each object MUST follow this exact schema:
    {{
        "title": "Slide Title",
        "layout_type": "top_image_bottom_cards",
        "theme_variant": "ocean",
        "visual_priority": "balanced",
        "background_design": {{
            "base_fill": "#F4F8FB",
            "elements": [
                {{
                    "type": "circle",
                    "x": -0.6,
                    "y": -0.9,
                    "w": 3.0,
                    "h": 3.0,
                    "fill": "#BFE3FF",
                    "transparency": 0.18
                }},
                {{
                    "type": "rounded_rect",
                    "x": 0.7,
                    "y": 1.3,
                    "w": 11.2,
                    "h": 4.9,
                    "fill": "#EAF3FF",
                    "transparency": 0.24
                }}
            ]
        }},
        "background": {{
            "has_background_image": true,
            "background_image_prompt": "Dark, moody abstract geometric corporate background, deep blue and charcoal, 8k resolution"
        }},
        "foreground_image": {{
            "has_foreground_image": true,
            "foreground_image_prompt": "Crisp 3D isometric render of a glowing data server rack, transparent background"
        }},
        "existing_chart": {{
            "has_existing_chart": true,
            "chart_topic_reference": "Topic"
        }},
        "content_blocks": [
            {{
                "block_title": "Efficiency",
                "body": "Short supporting sentence expanding the main idea.",
                "bullets": ["Increased speed by 40%", "Reduced server load"]
            }},
            {{
                "block_title": "Security",
                "body": "",
                "bullets": ["End-to-end encryption", "Zero-trust architecture"]
            }}
        ]
    }}

    Raw slide data to process:
    {summarized_json_string}
    """

    if progress_callback:
        progress_callback("Art Director is designing premium magazine layouts...", 75)

    response = client.models.generate_content(
        model="gemini-3.1-flash-lite-preview", 
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.5, 
        ),
    )
    
    return response.text

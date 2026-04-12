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
    You are an expert Presentation Art Director designing a high-end corporate pitch deck.
    I will give you a JSON array of raw slide data. Your job is to assign a complex, highly visual layout grid to each slide, group the text into logical blocks, and write prompts for AI image generation.

    make visually stunning and unique layout: have some examples but preffer variation across slide layouts:
    1. "top_image_bottom_cards" (Upper 50% is a striking foreground image. Lower 50% contains 2 to 3 side-by-side text boxes)
    2. "hero_background_with_floating_cards" (A full-bleed background image, with multiple text boxes floating on top of it)
    3. "split_left_image_right_cards" (Left 50% is an image. Right 50% contains multiple stacked text boxes)
    4. "diagonal_text_split" (Two separate text blocks placed diagonally. A foreground image goes in the center)
    5. "multi_column_cards" (No foreground image, just a subtle background image with 2 to 4 side-by-side vertical text boxes)

    CRITICAL INSTRUCTION FOR IMAGES & CHARTS:
    I already have real, highly-detailed data charts extracted for the following topics:
    {available_topics}
    
    If the incoming raw slide data contains an "existing_chart" object, you MUST preserve it in your final output. Do not delete it!

    add image prompt remember: images shoud be more abstract and artistic, like 3d designs and cartoonie. should not have realistic images They should evoke the theme of the slide rather than directly depict the charts.
    
    Return ONLY a valid JSON array of objects. 
    Each object MUST follow this exact schema:
    {{
        "title": "Slide Title",
        "layout_type": "top_image_bottom_cards",
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
                "bullets": ["Increased speed by 40%", "Reduced server load"]
            }},
            {{
                "block_title": "Security",
                "bullets": ["End-to-end encryption", "Zero-trust architecture"]
            }}
        ]
    }}

    Raw slide data to process:
    {summarized_json_string}
    """

    if progress_callback:
        progress_callback("Art Director is designing premium magazine layouts...", 75)

    # I updated the model name to match what you are using in your other file
    response = client.models.generate_content(
        model="gemini-3.1-flash-lite-preview", 
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.5, 
        ),
    )
    
    return response.text
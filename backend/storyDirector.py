import json
import os
import time
from typing import List
from dotenv import load_dotenv
from google import genai
from google.genai import types

load_dotenv()

CHUNK_CHAR_LIMIT = 300000 
CHUNK_OVERLAP = 5000
INTERMEDIATE_SLIDES_PER_CHUNK = 10
FINAL_SLIDE_LIMIT = 15

def _get_client() -> genai.Client:
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY is not set in the environment.")
    return genai.Client(api_key=api_key)

def _split_markdown_into_chunks(markdown_text: str) -> List[str]:
    text = markdown_text.strip()
    if not text: return []
    if len(text) <= CHUNK_CHAR_LIMIT: return [text]

    paragraphs = text.split("\n\n")
    chunks: List[str] = []
    current = ""

    for paragraph in paragraphs:
        candidate = f"{current}\n\n{paragraph}".strip() if current else paragraph
        if len(candidate) <= CHUNK_CHAR_LIMIT:
            current = candidate
            continue

        if current:
            chunks.append(current)
            overlap = current[-CHUNK_OVERLAP:] if len(current) > CHUNK_OVERLAP else current
            current = f"{overlap}\n\n{paragraph}".strip()
        else:
            start = 0
            while start < len(paragraph):
                end = start + CHUNK_CHAR_LIMIT
                chunks.append(paragraph[start:end])
                start = max(end - CHUNK_OVERLAP, start + 1)
            current = ""

        while len(current) > CHUNK_CHAR_LIMIT:
            chunks.append(current[:CHUNK_CHAR_LIMIT])
            current = current[CHUNK_CHAR_LIMIT - CHUNK_OVERLAP :].strip()

    if current:
        chunks.append(current)
    return chunks

def _generate_json_response(client: genai.Client, prompt: str) -> str:
    for attempt in range(3):
        try:
            response = client.models.generate_content(
                model="gemini-3.1-flash-lite-preview",       # gemini-3.1-Flash-Lite-preview is also an option if available
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.2,
                ),
            )
            return response.text
        except Exception as e:
            if attempt < 2:
                time.sleep(15) 
            else:
                raise ValueError(f"Gemini API Error after 3 tries: {str(e)}")

def _summarize_chunk(client: genai.Client, chunk_text: str, chunk_index: int, total_chunks: int) -> List[dict]:
    prompt = f"""You are analyzing chunk {chunk_index} of {total_chunks} from a very large markdown document.
Your goal is to create content for premium PowerPoint slides, not document pages.

CONTENT PRINCIPLES:
- Make each slide insight-led, not text-dump-led.
- Provide enough written substance so the slide can feel complete and persuasive.
- Favor structured thinking, frameworks, comparisons, implications, and key takeaways.
- When numbers exist, surface them explicitly so they can become visual callouts.
- Avoid vague bullets and avoid repeating the same point with different wording.

Return ONLY a valid JSON array.
Each object must have:
- "title": string
- "summary": 1 meaningful sentence, usually 18 to 34 words, that captures the slide's core message
- "bullets": array of 5 to 7 substantive bullet strings, each usually 12 to 26 words
- "highlight_metrics": array of 1 to 5 short metric/value strings pulled from the source when available
- "recommended_visual": one of "chart", "timeline", "framework", "comparison", "dashboard", "process", "map", "illustration"

Optional:
- "existing_chart": {{"has_existing_chart": true, "chart_topic_reference": "Topic"}}

CRITICAL:
If the text mentions a [CHART: Topic], [VISUALIZATION: Topic], [REFERENCE_CHART_EXTRACTED], or an image ![Topic](...), you MUST preserve that by including "existing_chart".
If the source contains numbers, dates, growth rates, percentages, budgets, market sizes, milestones, or phase-based plans, you must surface them in bullets and highlight_metrics.
Prefer slide-worthy content density over minimal summaries.

Limit to {INTERMEDIATE_SLIDES_PER_CHUNK} slide objects.
Markdown chunk:
{chunk_text}"""
    chunk_response = _generate_json_response(client, prompt)
    return json.loads(chunk_response)

def _merge_chunk_summaries(client: genai.Client, chunk_summaries: List[List[dict]]) -> List[dict]:
    serialized_summaries = json.dumps(chunk_summaries, ensure_ascii=True)
    prompt = f"""You are an expert presentation strategist.
Deduplicate and merge these slide candidates into a final presentation outline.

OUTPUT GOAL:
- Build a presentation that feels rich, structured, and executive-ready.
- Prefer slides with a clear key message, 4 to 6 strong supporting points, and surfaced metrics.
- Keep the story flowing from high-level context to analysis, then implications and recommendations.

Return ONLY a valid JSON array.
Each object must contain:
- "title"
- "summary"
- "bullets" (array of 5 to 7 strings)
- "highlight_metrics" (array of 1 to 5 strings when available)
- "recommended_visual" (one of "chart", "timeline", "framework", "comparison", "dashboard", "process", "map", "illustration")

CRITICAL:
- Preserve the "existing_chart" data if it exists in the candidates.
- Do not collapse everything into overly short bullets.
- Keep metrics explicit and presentation-ready.
- Prefer richer slides with real analytical content over sparse summaries.
- When the material includes steps, phases, or a roadmap, preserve that structure clearly.

Maximum {FINAL_SLIDE_LIMIT} slides. Order from high-level to specific.
Chunk slide candidates:
{serialized_summaries}"""
    merged_response = _generate_json_response(client, prompt)
    return json.loads(merged_response)

def generate_slide_content(markdown_text: str, progress_callback=None) -> str:
    client = _get_client()
    
    if progress_callback: progress_callback("Splitting document...", 5)
    chunks = _split_markdown_into_chunks(markdown_text)

    if not chunks: return "[]"

    chunk_summaries = []
    total_chunks = len(chunks)
    
    for index, chunk_text in enumerate(chunks, start=1):
        if progress_callback: 
            base_progress = 10 + ((index - 1) / total_chunks * 70) 
            progress_callback(f"Analyzing chunk {index} of {total_chunks}...", int(base_progress))
            
        chunk_summaries.append(_summarize_chunk(client, chunk_text, index, total_chunks))
        
        if index < total_chunks:
            if progress_callback:
                progress_callback(f"Cooling down API (protecting rate limits)...", int(base_progress + 5))
            time.sleep(60)

    if progress_callback: progress_callback("Merging final presentation...", 85)
    final_slides = _merge_chunk_summaries(client, chunk_summaries)
    
    if progress_callback: progress_callback("AI Generation Complete!", 100)
    return json.dumps(final_slides, ensure_ascii=True)

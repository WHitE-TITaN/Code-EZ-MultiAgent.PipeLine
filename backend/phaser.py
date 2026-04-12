import json
import re

def parse_markdown_to_json(markdown_text):
    """
    Parses raw markdown text into slides AND harvests the raw Base64 image data.
    """
    slides = []
    harvested_images = {} # THE HARVESTER DICTIONARY
    
    current_slide = {
        "title": "Introduction",
        "content": [],
        "raw_text": "",
        "existing_chart": {} 
    }
    
    lines = markdown_text.split('\n')
    
    for line in lines:
        clean_line = line.strip()
        
        if not clean_line:
            continue
            
        # 1. New Slide Trigger
        if clean_line.startswith('# ') or clean_line.startswith('## '):
            if current_slide["content"] or current_slide["title"] != "Introduction" or current_slide.get("existing_chart"):
                slides.append(current_slide)
                
            title_text = clean_line.lstrip('#').strip()
            current_slide = {
                "title": title_text,
                "content": [],
                "raw_text": "",
                "existing_chart": {}
            }
            
        # 2. THE VISUALIZATION SEPARATOR LOGIC
        elif clean_line.upper().startswith('[CHART:') or clean_line.upper().startswith('[VISUALIZATION:'):
            topic = re.sub(r'\[(CHART|VISUALIZATION):\s*', '', clean_line, flags=re.IGNORECASE).rstrip(']')
            current_slide["existing_chart"] = {
                "has_existing_chart": True, 
                "chart_topic_reference": topic
            }
            current_slide["raw_text"] += f"{clean_line}\n"

        # 3. HARVEST THE BASE64 IMAGES!
        elif clean_line.startswith('!['):
            # NEW REGEX: Grabs the Topic (group 1) AND the Base64 URL (group 2)
            match = re.search(r'!\[(.*?)\]\((.*?)\)', clean_line)
            if match:
                topic = match.group(1).strip()
                image_data = match.group(2).strip() # The massive Base64 string
                
                # Tell the slide it has a chart
                current_slide["existing_chart"] = {
                    "has_existing_chart": True, 
                    "chart_topic_reference": topic
                }
                
                # Save the massive string to our dictionary!
                if topic not in harvested_images:
                    harvested_images[topic] = []
                harvested_images[topic].append(image_data)
                
            current_slide["raw_text"] += f"{clean_line}\n"
            
        # 4. Bullets
        elif clean_line.startswith('- ') or clean_line.startswith('* '):
            bullet_text = clean_line[2:].strip()
            current_slide["content"].append({"type": "bullet", "text": bullet_text})
            current_slide["raw_text"] += f"{clean_line}\n"
            
        # 5. Standard Paragraphs
        else:
            current_slide["content"].append({"type": "paragraph", "text": clean_line})
            current_slide["raw_text"] += f"{clean_line}\n"
            
    if current_slide["content"] or current_slide.get("existing_chart"):
        slides.append(current_slide)
        
    # We now return a tuple containing BOTH the slides and the images!
    return slides, harvested_images
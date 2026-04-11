import json


def parse_markdown_to_json(markdown_text):
    """
    Parses raw markdown text into a structured list of slide dictionaries.
    """
    slides = []
    current_slide = {
        "title": "Introduction", # Default title if no heading is found first
        "content": [],
        "raw_text": ""
    }
    
    # Split the text into individual lines
    lines = markdown_text.split('\n')
    
    for line in lines:
        clean_line = line.strip()
        
        # Skip completely empty lines to keep data clean
        if not clean_line:
            continue
            
        # Check if the line is a  Heading 1 or Heading 2 (New Slide trigger)
        if clean_line.startswith('# ') or clean_line.startswith('## '):
            # 1. Save the slide we were just working on (if it has data)
            if current_slide["content"] or current_slide["title"] != "Introduction":
                slides.append(current_slide)
                
            # 2. Start a brand new slide
            # .lstrip('#').strip() removes the markdown hashes to get pure text
            title_text = clean_line.lstrip('#').strip()
            current_slide = {
                "title": title_text,
                "content": [],
                "raw_text": ""
            }
            
        # Check if it's a bullet point
        elif clean_line.startswith('- ') or clean_line.startswith('* '):
            bullet_text = clean_line[2:].strip()
            current_slide["content"].append({"type": "bullet", "text": bullet_text})
            current_slide["raw_text"] += f"{clean_line}\n"
            
        # Otherwise, treat it as a normal paragraph of text
        else:
            current_slide["content"].append({"type": "paragraph", "text": clean_line})
            current_slide["raw_text"] += f"{clean_line}\n"
            
    # Don't forget to append the very last slide when the loop finishes!
    if current_slide["content"]:
        slides.append(current_slide)


    
        
    return slides
import re

def extract_and_map_visualizations(markdown_text: str):
    """
    Reads markdown, tracks the current heading (topic), 
    extracts base64 images into a dictionary, and returns the clean text.
    """
    visualizations_by_topic = {}
    current_topic = "General Overview" # Default topic if no header is found first
    
    cleaned_lines = []
    lines = markdown_text.split('\n')
    
    # Regex to capture the full base64 data string: (data:image/png;base64,...)
    b64_pattern = re.compile(r"\(data:image\/[a-zA-Z0-9+.-]+;base64,[^\)]+\)")

    for line in lines:
        # 1. Update the active topic if we hit a Markdown heading (#, ##, ###)
        if line.strip().startswith('#'):
            # Strip the hashes and spaces to get the clean topic name
            current_topic = line.lstrip('#').strip()
            
        # 2. Look for Base64 image data in the current line
        matches = b64_pattern.findall(line)
        
        if matches:
            # Create the list for this topic if it doesn't exist
            if current_topic not in visualizations_by_topic:
                visualizations_by_topic[current_topic] = []
                
            # Store the extracted image data
            visualizations_by_topic[current_topic].extend(matches)
            
            # 3. Strip the massive data from the line, but leave a breadcrumb for the AI
            # This tells the summarizer "Hey, a chart used to be here!"
            clean_line = b64_pattern.sub("[REFERENCE_CHART_EXTRACTED]", line)
            clean_line = clean_line.replace("Visualization[]", "").replace("![]", "").strip()
            
            if clean_line:
                cleaned_lines.append(clean_line)
        else:
            # If it's just an empty formatting tag left over, skip it
            if "Visualization[]" in line.strip():
                continue
            cleaned_lines.append(line)

    # Rejoin the text for the Summarizer AI
    clean_markdown = "\n".join(cleaned_lines)
    
    return clean_markdown, visualizations_by_topic
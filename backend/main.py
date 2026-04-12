from dotenv import load_dotenv
load_dotenv()

import uuid
import os
from fastapi import FastAPI, UploadFile, File, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse # ADDED THIS

from backend.phaser import extract_and_map_visualizations
from backend.storyDirector import generate_slide_content
from backend.compiler import build_powerpoint # ADDED THIS (adjust import path if needed)
from backend.artDirector import design_slide_layouts
import json # You will need this to parse the string into a dictionary

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], 
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

job_statuses = {}

def process_pipeline(job_id: str, markdown_text: str):
    def update_progress(message: str, percent: int):
        job_statuses[job_id]["message"] = message
        job_statuses[job_id]["progress"] = percent
        print(f"[{job_id}] {percent}% - {message}")

    try:
        # --- AGENT 1: The Analyst ---
        update_progress("Parsing markdown structure...", 10)
        
        # CAUGHT! We unpack both the slides and the image dictionary here
        slides_data, extracted_images = extract_and_map_visualizations(markdown_text) 
        
        original_size = len(markdown_text)
        summarized_size = len(slides_data)
        print(f"📊 TEXT COMPRESSION: Original ({original_size} chars) -> Summarized ({summarized_size} chars)")

        update_progress("Generating presentation content...", 20)
        summarized_slides_string = generate_slide_content(slides_data, progress_callback=update_progress)
        
        # --- AGENT 2: The Art Director ---
        update_progress("Designing premium layouts and image prompts...", 50)
        
        # PASS THE KEYS to the Art Director so it knows what charts we saved!
        available_topics = list(extracted_images.keys())
        final_blueprint_string = design_slide_layouts(summarized_slides_string, available_topics, progress_callback=update_progress)
        
        blueprint_data = json.loads(final_blueprint_string)

        # --- AGENT 3: The Compiler ---
        update_progress("Compiling PowerPoint and rendering images...", 80)
        output_filename = f"Deck_{job_id}.pptx"
        
        # PASS THE DICTIONARY to the compiler so it can draw them!
        final_file_path = build_powerpoint(blueprint_data, extracted_images, output_filename=output_filename)

        # Final success state
        job_statuses[job_id]["status"] = "completed"
        job_statuses[job_id]["message"] = "Presentation Ready for Download!"
        job_statuses[job_id]["progress"] = 100
        job_statuses[job_id]["file_path"] = output_filename 
        
    except Exception as e:
        job_statuses[job_id]["status"] = "error"
        job_statuses[job_id]["message"] = f"Pipeline failed: {str(e)}"
        print(f"Pipeline crashed: {str(e)}")

@app.get("/")
def read_root():
    return {"message": "Markdown to PPTX Converter API"}

@app.post("/upload")
async def upload_markdown(file: UploadFile = File(...), background_tasks: BackgroundTasks = BackgroundTasks()):
    try:
        content = await file.read()
        markdown_text = content.decode('utf-8')
        
        job_id = str(uuid.uuid4())
        
        job_statuses[job_id] = {
            "status": "processing",
            "message": "File uploaded successfully. Starting AI agents...",
            "progress": 0,
        }
        
        background_tasks.add_task(process_pipeline, job_id, markdown_text)

        return {
            "status": "success",
            "job_id": job_id,
            "message": "Background processing started"
        }
    except Exception as e:
        return {"status": "error", "message": str(e)}

@app.get("/status/{job_id}")
def get_job_status(job_id: str):
    if job_id not in job_statuses:
        return {"status": "error", "message": "Job ID not found"}
    return job_statuses[job_id]

# --- ADDED THIS NEW ENDPOINT TO DOWNLOAD THE FILE ---
@app.get("/download/{job_id}")
def download_presentation(job_id: str):
    if job_id not in job_statuses or job_statuses[job_id]["status"] != "completed":
        return {"status": "error", "message": "File not ready or job not found"}
    
    file_path = job_statuses[job_id]["file_path"]
    
    if os.path.exists(file_path):
        # Serve the file to the browser
        return FileResponse(path=file_path, filename="AI_Presentation.pptx", media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation')
    else:
        return {"status": "error", "message": "File was generated but is missing from the server"}
from dotenv import load_dotenv
load_dotenv()

from fastapi import FastAPI, UploadFile, File, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from backend.phaser import parse_markdown_to_json
from backend.storyDirector import generate_slide_content
import uuid

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], 
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory dictionary to track jobs
job_statuses = {}

# --- THE BACKGROUND WORKER ---
def process_pipeline(job_id: str, markdown_text: str):
    # This callback updates the dictionary, which the frontend reads
    def update_progress(message: str, percent: int):
        job_statuses[job_id]["message"] = message
        job_statuses[job_id]["progress"] = percent
        print(f"[{job_id}] {percent}% - {message}")

    try:
        update_progress("Parsing markdown structure...", 2)
        slides_data = parse_markdown_to_json(markdown_text)
        
        # Pass the callback into the summarizer!
        summarized_slides_json = generate_slide_content(markdown_text, progress_callback=update_progress)
        print({summarized_slides_json})

        print (f"[{job_id}] Final summarized slides: {summarized_slides_json}")
        

        # Final success state
        job_statuses[job_id]["status"] = "completed"
        job_statuses[job_id]["message"] = "Presentation Data Ready!"
        job_statuses[job_id]["progress"] = 100
        job_statuses[job_id]["data"] = summarized_slides_json
        
    except Exception as e:
        job_statuses[job_id]["status"] = "error"
        job_statuses[job_id]["message"] = f"Pipeline failed: {str(e)}"




@app.get("/")
def read_root():
    return {"message": "Markdown to PPTX Converter API"}

@app.get("/health")
def health_check():
    return {"status": "ok"}

@app.post("/upload")
async def upload_markdown(file: UploadFile = File(...), background_tasks: BackgroundTasks = BackgroundTasks()):
    try:
        content = await file.read()
        markdown_text = content.decode('utf-8')
        
        # Create a unique ticket for this process
        job_id = str(uuid.uuid4())
        
        # Initialize status
        job_statuses[job_id] = {
            "status": "processing",
            "message": "File uploaded successfully. Starting AI agents...",
            "progress": 0,
            "data": None
        }
        
        # Hand off to the background task (does NOT block the response)
        background_tasks.add_task(process_pipeline, job_id, markdown_text)

        # Return the ticket immediately
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
from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    # For development, "*" allows all domains. 
    # Before the hackathon ends, change "*" to your actual Vercel URL!
    allow_origins=["*"], 
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def read_root():
    return {"message": "Markdown to PPTX Converter API"}

@app.get("/health")
def health_check():
    return {"status": "ok"}

@app.post("/upload")
async def upload_markdown(file: UploadFile = File(...)):
    """
    Receive a markdown file and prepare it for PPTX conversion
    """
    try:
        # Read file content
        content = await file.read()
        
        # Decode content
        markdown_text = content.decode('utf-8')
        
        print(f"Received file: {file.filename}")
        print(f"File size: {len(markdown_text)} characters")
        print(f"Content preview: {markdown_text[:200]}...")
        
        # For now, just acknowledge receipt
        # TODO: Add markdown parsing and PPTX generation logic here
        return {
            "status": "success",
            "message": f"File {file.filename} received successfully",
            "filename": file.filename,
            "size": len(markdown_text)
        }
    
    except Exception as e:
        print(f"Error uploading file: {str(e)}")
        return {
            "status": "error",
            "message": str(e)
        }

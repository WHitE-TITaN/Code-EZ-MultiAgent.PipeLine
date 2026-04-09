from fastapi import FastAPI
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
    print("hello world")
    return {"message": "hello world"}

@app.get("/health")
def health_check():
    return {"status": "ok"}
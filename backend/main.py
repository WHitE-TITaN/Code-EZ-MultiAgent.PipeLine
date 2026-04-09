from fastapi import FastAPI

app = FastAPI()

@app.get("/")
def read_root():
    print("hello world")
    return {"message": "hello world"}

@app.get("/health")
def health_check():
    return {"status": "ok"}
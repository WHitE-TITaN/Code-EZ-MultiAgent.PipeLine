# Code-EZ MultiAgent Pipeline

AI-powered Markdown-to-PowerPoint pipeline that turns long-form research documents into consulting-style presentation decks.

The project combines a `Next.js` upload UI with a `FastAPI` backend that runs a three-stage generation flow:

1. Parse the Markdown and extract embedded chart images.
2. Summarize the content into a slide storyline with Gemini.
3. Design slide blueprints and compile a `.pptx` deck.

## What It Does

- Accepts `.md` files from the web UI.
- Processes jobs asynchronously in the backend.
- Tracks job progress with polling.
- Extracts base64 charts/images from Markdown and maps them to section topics.
- Uses Gemini to summarize large documents into slide-ready content.
- Uses a second Gemini pass to assign premium slide layouts, themes, and image directions.
- Compiles a downloadable PowerPoint deck with `python-pptx`.
- Includes sample Markdown inputs and generated decks under `tests/`.

## Architecture Diagram

```mermaid
flowchart LR
    U["User"] --> F["Frontend<br/>Next.js 16 + React 19"]
    F -->|POST /upload| B["FastAPI Backend"]
    F -->|GET /status/{job_id}| B
    F -->|GET /download/{job_id}| B

    B --> P["Pipeline Orchestrator<br/>backend/main.py"]
    P --> A1["Agent 1: Phaser<br/>Extract headings + base64 charts"]
    P --> A2["Agent 2: Story Director<br/>Chunk, summarize, merge slides"]
    P --> A3["Agent 3: Art Director<br/>Assign layouts, themes, visuals"]
    P --> C["Compiler<br/>python-pptx deck builder"]

    A2 --> G["Google Gemini API"]
    A3 --> G
    A1 --> X["Extracted chart/topic map"]
    X --> C
    A3 --> C
    C --> O["Generated .pptx file<br/>saved on server"]
    O --> B
```

## Pipeline Flow

### 1. Upload and job creation
- The frontend accepts a Markdown file and sends it to `POST /upload`.
- The backend creates a `job_id`, stores an in-memory status record, and starts background processing.

### 2. Phaser
- File: `backend/phaser.py`
- Scans headings to understand topic boundaries.
- Detects inline base64 images like `data:image/...;base64,...`.
- Replaces large image blobs with `[REFERENCE_CHART_EXTRACTED]` markers so the LLM can keep chart context without huge payloads.
- Returns:
  - cleaned Markdown
  - topic-to-images mapping

### 3. Story Director
- File: `backend/storyDirector.py`
- Splits large Markdown into chunks.
- Sends each chunk to Gemini for slide candidate generation.
- Merges chunk outputs into a final presentation outline.
- Preserves chart references when the source implies an existing visualization.
- Emits structured slide content such as:
  - title
  - summary
  - bullets
  - highlight metrics
  - recommended visual type

### 4. Art Director
- File: `backend/artDirector.py`
- Takes summarized slide JSON and transforms it into a visual blueprint.
- Chooses:
  - layout type
  - theme variant
  - visual priority
  - background design elements
  - content blocks
  - whether to use extracted charts or generated visuals

### 5. Compiler
- File: `backend/compiler.py`
- Builds the final deck with `python-pptx`.
- Adds a cover slide.
- Renders themed layouts and decorative elements.
- Places extracted charts when topic matches are found.
- Generates fallback visual assets for slides when needed.
- Saves the final file as `Deck_<job_id>.pptx`.

### 6. Polling and download
- The frontend polls `GET /status/{job_id}` every 2 seconds.
- Once complete, it redirects the browser to `GET /download/{job_id}` to fetch the PowerPoint.

## Tech Stack

### Frontend
- Next.js 16
- React 19
- TypeScript
- Tailwind CSS 4

### Backend
- FastAPI
- Uvicorn
- Google GenAI SDK
- python-pptx
- Pillow
- requests
- python-dotenv

### Deployment
- Docker
- GitHub Actions
- Hugging Face Spaces sync workflow

## Project Structure

```text
Code-EZ-MultiAgent.PipeLine/
|-- backend/
|   |-- main.py              # FastAPI app and pipeline orchestration
|   |-- phaser.py            # Markdown cleanup and chart extraction
|   |-- storyDirector.py     # Slide summarization and merge logic
|   |-- artDirector.py       # Layout/theme blueprint generation
|   |-- compiler.py          # PPTX generation
|   |-- requirements.txt     # Python dependencies
|   `-- .env                 # Backend environment variables
|-- frontend/
|   |-- app/
|   |   |-- page.tsx         # Upload UI, polling, download redirect
|   |   |-- layout.tsx       # App shell metadata
|   |   `-- globals.css      # Global styles
|   |-- package.json
|   `-- next.config.ts
|-- tests/                   # Sample markdown inputs and generated outputs
|-- .github/workflows/
|   `-- synTo_HF.yml         # Pushes repo to Hugging Face Space
|-- Dockerfile               # Backend container image
`-- README.md
```

## API Endpoints

### `GET /`
Returns a basic API message.

### `POST /upload`
Uploads a Markdown file and starts background processing.

Request:

```bash
curl -X POST -F "file=@sample.md" http://localhost:7860/upload
```

Example response:

```json
{
  "status": "success",
  "job_id": "uuid-value",
  "message": "Background processing started"
}
```

### `GET /status/{job_id}`
Returns progress for a running or completed job.

Example response:

```json
{
  "status": "processing",
  "message": "Generating presentation content...",
  "progress": 20
}
```

Completed response includes `file_path`.

### `GET /download/{job_id}`
Downloads the generated PowerPoint when the job is complete.

## Local Development

### Prerequisites
- Node.js 18+
- Python 3.11+
- A valid `GEMINI_API_KEY`

### 1. Backend setup

```bash
cd backend
pip install -r requirements.txt
uvicorn main:app --host 0.0.0.0 --port 7860 --reload
```

For package-style imports from the repo root, this also works:

```bash
uvicorn backend.main:app --host 0.0.0.0 --port 7860 --reload
```

### 2. Frontend setup

```bash
cd frontend
npm install
npm run dev
```

### 3. Open the app

- Frontend: `http://localhost:3000`
- Backend: `http://localhost:7860`

## Environment Variables

The backend expects:

```env
GEMINI_API_KEY=your_api_key_here
```

Optional runtime values commonly used in deployment:

- `PORT`
- `HOST`
- `WORKERS`

## Docker

Build and run:

```bash
docker build -t code-ez-pipeline .
docker run -p 7860:7860 --env-file backend/.env code-ez-pipeline
```

The current Docker image runs the backend with multiple Uvicorn workers.

## Sample Inputs and Outputs

The `tests/` folder includes:

- sample Markdown research documents
- generated PowerPoint decks
- template decks / slide masters

These examples are useful for:

- evaluating slide quality
- checking formatting behavior
- validating chart extraction cases

## Key Implementation Notes

- Job state is stored in memory, so status is lost if the backend restarts.
- Generated `.pptx` files are written to the server filesystem, currently in the repo/runtime working directory.
- The frontend currently calls `http://localhost:7860` directly, so it is tied to local backend development unless configured differently.
- CORS is currently open to all origins.
- The Gemini-based summarization step adds cooldown delays between chunks to reduce rate-limit issues.
- The compiler can use extracted charts from the source Markdown and combine them with generated visual styling.
- Large documents are chunked before summarization to stay within model constraints.

## Current Limitations

- No persistent database or job queue.
- No authentication or multi-tenant isolation.
- No configurable frontend API base URL yet.
- No formal automated test suite for the API/UI flow.
- Dependency files may need tightening for fully reproducible fresh installs.
- Progress and output tracking are process-local, which matters when running multiple workers.

## Deployment Notes

- The repo includes a GitHub Actions workflow that force-pushes the `Master` branch to a Hugging Face Space.
- The root `Dockerfile` is backend-focused and does not build or serve the Next.js frontend.
- If you want a single deployed app, you will likely need:
  - frontend deployment separately
  - or a unified deployment strategy that serves both frontend and backend together


## License

This project appears to be a hackathon/prototype codebase.
---
title: MultiAgent PPTX Backend
emoji: 🚀
colorFrom: blue
colorTo: indigo
sdk: docker
pinned: false
---

# 📊 Markdown to PPTX Converter

> **Code EZ: Master of Agents Hackathon Submission**

A robust, scalable web application that converts Markdown files into structured, visually appealing PowerPoint presentations. Built with multi-user concurrent processing capabilities to handle multiple users simultaneously.

## 🌟 Features

### ✅ Core Functionality
- **File Upload**: Drag-and-drop interface for `.md` files
- **Real-time Processing**: Instant conversion feedback
- **Structured Output**: Well-formatted PPTX presentations
- **Error Handling**: Comprehensive validation and error messages

### 🚀 Advanced Capabilities
- **Multi-User Support**: Concurrent processing for multiple users
- **Scalable Architecture**: Multiple worker processes handle requests independently
- **Isolated Processing**: Each user request processed in separate worker instance
- **Load Balancing**: Automatic distribution of requests across workers
- **Fault Tolerance**: System continues operating if individual workers fail

### 🛠️ Technical Stack
- **Frontend**: Next.js 16, React 19, TypeScript, Tailwind CSS
- **Backend**: FastAPI, Python 3.11, Uvicorn
- **Deployment**: Docker, Hugging Face Spaces
- **Scaling**: Multiple worker processes (4 concurrent workers)

## 🏗️ Architecture

```
┌─────────────────┐    ┌─────────────────┐
│   Frontend      │    │   Backend       │
│   (Next.js)     │◄──►│   (FastAPI)     │
│                 │    │                 │
│ • File Upload   │    │ • File Receive  │
│ • Progress UI   │    │ • MD Processing │
│ • Download PPTX │    │ • PPTX Generate │
└─────────────────┘    └─────────────────┘
         │                       │
         └───────────────────────┘
              Worker Pool
          (4 concurrent processes)
```

## 🚀 Quick Start

### Prerequisites
- Node.js 18+
- Python 3.11+
- Docker (for deployment)

### Local Development

#### 1. Clone Repository
```bash
git clone https://github.com/WHitE-TITaN/Code-EZ-MultiAgent.PipeLine.git
cd Code-EZ-MultiAgent.PipeLine
```

#### 2. Backend Setup
```bash
cd backend
pip install -r requirements.txt
uvicorn main:app --host 0.0.0.0 --port 7860 --workers 4
```

#### 3. Frontend Setup
```bash
cd frontend
npm install
npm run dev
```

#### 4. Access Application
- **Frontend**: http://localhost:3000
- **Backend API**: http://localhost:7860
- **API Docs**: http://localhost:7860/docs

## 📡 API Documentation

### Endpoints

#### `GET /`
Returns API information
```json
{
  "message": "Markdown to PPTX Converter API"
}
```

#### `GET /health`
Health check endpoint
```json
{
  "status": "ok"
}
```

#### `POST /upload`
Upload markdown file for conversion
- **Content-Type**: `multipart/form-data`
- **Parameter**: `file` (markdown file)
- **Response**:
```json
{
  "status": "success",
  "message": "File test.md received successfully",
  "filename": "test.md",
  "size": 1234
}
```

## 🔄 Multi-User Concurrent Processing

### How It Works
The application uses **multiple worker processes** to handle concurrent requests:

1. **Load Balancing**: Uvicorn automatically distributes incoming requests across 4 worker processes
2. **Isolation**: Each user request is processed in a separate worker instance
3. **Concurrency**: Multiple users can upload and convert files simultaneously
4. **Fault Tolerance**: If one worker fails, others continue processing requests

### Scaling Configuration

#### Development (with hot reload)
```bash
uvicorn main:app --host 0.0.0.0 --port 7860 --reload
# Note: --workers ignored when --reload is enabled
```

#### Production (multiple workers)
```bash
uvicorn main:app --host 0.0.0.0 --port 7860 --workers 4
```

#### Docker Deployment
```dockerfile
CMD ["uvicorn", "backend.main:app", "--host", "0.0.0.0", "--port", "7860", "--workers", "4"]
```

## 🐳 Docker Deployment

### Build and Run
```bash
# Build image
docker build -t md-to-pptx .

# Run with multiple workers
docker run -p 7860:7860 md-to-pptx
```

### Hugging Face Spaces
The application is configured for deployment on Hugging Face Spaces with:
- Automatic CI/CD via GitHub Actions
- Docker containerization
- Multi-worker scaling
- CORS enabled for web access

## 📁 Project Structure

```
Code-EZ-MultiAgent.PipeLine/
├── backend/
│   ├── main.py              # FastAPI application
│   ├── requirements.txt     # Python dependencies
│   └── Dockerfile          # Container configuration
├── frontend/
│   ├── app/
│   │   ├── layout.tsx      # Root layout
│   │   ├── page.tsx        # Main upload interface
│   │   └── globals.css     # Global styles
│   ├── package.json        # Node dependencies
│   └── next.config.ts      # Next.js configuration
├── .github/
│   └── workflows/
│       └── synTo_HF.yml    # CI/CD pipeline
├── dockerfile              # Root Docker config
└── README.md              # This file
```

## 🧪 Testing

### Manual Testing
1. Start both frontend and backend
2. Open http://localhost:3000
3. Upload a `.md` file
4. Verify processing in backend logs
5. Test with multiple browser tabs for concurrent processing

### API Testing
```bash
# Health check
curl http://localhost:7860/health

# File upload test
curl -X POST -F "file=@test.md" http://localhost:7860/upload
```

## 🔧 Configuration

### Environment Variables
- `PORT`: Server port (default: 7860)
- `WORKERS`: Number of worker processes (default: 4)
- `HOST`: Server host (default: 0.0.0.0)

### CORS Settings
Currently configured for development with `allow_origins=["*"]`. Update for production deployment.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

This project is part of the Code EZ: Master of Agents hackathon submission.

## 🎯 Future Enhancements

- [ ] PPTX generation with custom slide masters
- [ ] Advanced markdown parsing (tables, code blocks, images)
- [ ] User authentication and file management
- [ ] Batch processing capabilities
- [ ] Custom theme support
- [ ] Export to multiple formats (PDF, HTML)

---

**Built with ❤️ for the Code EZ: Master of Agents Hackathon**

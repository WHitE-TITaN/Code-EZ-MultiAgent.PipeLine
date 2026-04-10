# Use an official Python runtime
FROM python:3.11-slim

# Set the working directory in the container
WORKDIR /app

# Copy backend requirements first for better Docker layer caching
COPY ./backend/requirements.txt .

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy the backend application code
COPY ./backend /app/backend

# Ensure backend is a Python package
RUN touch /app/backend/__init__.py

# Expose the port Hugging Face Spaces expects
EXPOSE 7860

# Run the FastAPI server with multiple workers for concurrent request handling
CMD ["uvicorn", "backend.main:app", "--host", "0.0.0.0", "--port", "7860", "--workers", "2"]

# Health check - verifies server is responding correctly
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
  CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:7860/health').read()"
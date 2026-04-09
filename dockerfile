# Use an official Python runtime
FROM python:3.11-slim

# Set the working directory in the container
WORKDIR /app

# Copy JUST the backend requirements first (better caching)
COPY ./backend/requirements.txt /app/requirements.txt

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the backend code into the container
COPY ./backend /app/backend

# Expose the port Hugging Face expects
EXPOSE 7860

# Run the FastAPI server (pointing to backend/main.py)
CMD ["uvicorn", "backend.main:app", "--host", "0.0.0.0", "--port", "7860"]
FROM python:3.11-slim

WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files and dataset
COPY . .

# Hugging Face Spaces exposes port 7860
EXPOSE 7860

# Start WSGI server
CMD ["gunicorn", "-b", "0.0.0.0:7860", "--workers=2", "--timeout=120", "app:server"]

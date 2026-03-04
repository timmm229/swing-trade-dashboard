FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p output

EXPOSE 5000

# The PORT env var is used by Railway, Render, Fly.io, etc.
CMD ["python", "app.py"]

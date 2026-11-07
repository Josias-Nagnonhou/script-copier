FROM python:3.12-slim

RUN apt-get update && apt-get install -y --no-install-recommends ffmpeg ca-certificates fontconfig fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*
# Deno : nécessaire à yt-dlp pour lire YouTube
COPY --from=denoland/deno:bin /deno /usr/local/bin/deno

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -U -r requirements.txt
COPY . .

ENV PYTHONUNBUFFERED=1 MALLOC_ARENA_MAX=2
CMD gunicorn -w 1 --threads 4 --timeout 120 -b 0.0.0.0:${PORT:-10000} app:app

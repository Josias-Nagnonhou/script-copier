FROM python:3.12-slim

RUN apt-get update && apt-get install -y --no-install-recommends ffmpeg ca-certificates fontconfig fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*
# Deno : nécessaire à yt-dlp pour lire YouTube
COPY --from=denoland/deno:bin /deno /usr/local/bin/deno
# DeepFilterNet : retire la musique et les bruits de fond pour garder la voix (onglet Nettoyer)
RUN python -c "import urllib.request as u; u.urlretrieve('https://github.com/Rikorose/DeepFilterNet/releases/download/v0.5.6/deep-filter-0.5.6-x86_64-unknown-linux-musl', '/usr/local/bin/deep-filter')" \
    && chmod 755 /usr/local/bin/deep-filter && /usr/local/bin/deep-filter --version

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -U -r requirements.txt
# Détourage des personnes pour les miniatures (MediaPipe, léger) : le visage n'est jamais redessiné
RUN pip install --no-cache-dir --no-deps mediapipe==0.10.32 && pip install --no-cache-dir absl-py flatbuffers matplotlib \
    && mkdir -p /opt/models && python -c "import urllib.request as u; u.urlretrieve('https://storage.googleapis.com/mediapipe-models/image_segmenter/selfie_multiclass_256x256/float32/latest/selfie_multiclass_256x256.tflite', '/opt/models/selfie_multiclass_256x256.tflite')" \
    && python -c "import mediapipe; from mediapipe.tasks.python import vision; print('mediapipe ok', mediapipe.__version__)"
ENV THUMB_SEG_MODEL=/opt/models/selfie_multiclass_256x256.tflite
COPY . .

ENV PYTHONUNBUFFERED=1 MALLOC_ARENA_MAX=2
CMD gunicorn -w 1 --threads 4 --timeout 120 -b 0.0.0.0:${PORT:-10000} app:app

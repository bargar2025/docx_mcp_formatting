# syntax=docker/dockerfile:1
FROM python:3.11-slim AS base

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PORT=3001

# System deps for Pillow image formats
RUN apt-get update && apt-get install -y --no-install-recommends \
    libjpeg62-turbo \
    zlib1g \
    libpng16-16 \
    libopenjp2-7 \
    libtiff6 \
    libfreetype6 \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install runtime deps (no build backend required)
RUN pip install --no-cache-dir \
    mcp==1.4.1 \
    python-docx>=1.1.0 \
    azure-storage-blob>=12.19.0 \
    Pillow>=10.0.0

# Copy source
COPY src ./src
COPY pyproject.toml README.md USAGE_EXAMPLES.md IMPLEMENTATION_SUMMARY.md ./

# Create non-root user
RUN useradd -m appuser
USER appuser

EXPOSE 3001

# Start FastMCP over SSE, bound to 0.0.0.0 for ACA ingress
CMD ["python", "-m", "src", "sse"]
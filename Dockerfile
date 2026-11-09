# Multi-stage Dockerfile for excel-mcp (luiccian-excel-mcp)
FROM ghcr.io/astral-sh/uv:0.6.14 AS uv-bin

FROM python:3.12-slim

# Copy uv binary
COPY --from=uv-bin /uv /uvx /bin/

# Environment variables
ENV PYTHONUNBUFFERED=1 \
    UV_COMPILE_BYTECODE=1 \
    UV_LINK_MODE=copy \
    MCP_ALLOWED_DIRECTORIES=/data

WORKDIR /app

# Create a non-root application user and data directory
RUN useradd -m -u 10001 appuser && \
    mkdir -p /data && \
    chown -R appuser:appuser /data /app

# Copy dependency definition files
COPY pyproject.toml README.md LICENSE /app/

# Copy application source code
COPY src/ /app/src/

# Install application and production dependencies into virtual environment
RUN uv sync --no-dev --no-editable

# Switch to non-root user
USER appuser

ENV PATH="/app/.venv/bin:$PATH"

# Run the excel-mcp server
ENTRYPOINT ["excel-mcp"]
CMD ["server"]

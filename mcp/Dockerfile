FROM python:3.11-alpine

# Install system dependencies needed by Java compilation and Python runtime
RUN apk add --no-cache openjdk11-jdk bash gcc musl-dev libffi-dev

# Repository root in the image
WORKDIR /app

# Copy both Python server code and Java implementation into the image
COPY mcp/ ./mcp/
COPY java/ ./java/

# Install uv and Python dependencies (mcp/ contains pyproject.toml)
WORKDIR /app/mcp
RUN pip install --no-cache-dir uv && \
    uv pip install --system .

# Compile Java tools to /app/java/dist so the Python wrapper can load xelixir.* classes
WORKDIR /app/java
RUN bash ./tools/compile.sh

# Default entrypoint: run the Excel MCP server via the CLI entrypoint
# (so CLI args like `-t sse -p 8585` are parsed by the `xelixir` script)
WORKDIR /app/mcp
ENTRYPOINT ["uv", "run", "xelixir"]

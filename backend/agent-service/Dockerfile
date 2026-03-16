FROM python:3.11-slim

# Install Node.js for syntax validation
RUN apt-get update && \
    apt-get install -y --no-install-recommends nodejs npm curl && \
    npm install -g acorn && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Install Claude Code CLI (required by claude-agent-sdk)
RUN npm install -g @anthropic-ai/claude-code

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY agent_service/ agent_service/

# Create non-root user (Claude CLI refuses --dangerously-skip-permissions as root)
RUN useradd -m -s /bin/bash agent && chown -R agent:agent /app
USER agent

EXPOSE 8081

CMD ["uvicorn", "agent_service.main:app", "--host", "0.0.0.0", "--port", "8081"]

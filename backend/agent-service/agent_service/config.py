"""Configuration for the agent service."""
import os


ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
MODEL_LARGE = "claude-sonnet-4-6"
MODEL_SMALL = "claude-haiku-4-5-20251001"
SMALL_MODULE_THRESHOLD = 50  # lines
MAX_VALIDATION_RETRIES = 3
MAX_AGENT_TURNS = 10
MAX_CONCURRENCY = 5
PORT = int(os.getenv("AGENT_SERVICE_PORT", "8081"))

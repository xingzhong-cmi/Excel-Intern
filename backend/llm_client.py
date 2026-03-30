"""LLM API client module for Excel Assistant."""

import os
import logging
import httpx

logger = logging.getLogger(__name__)

# Default configuration
DEFAULT_API_URL = "https://api.deepseek.com/v1/chat/completions"
DEFAULT_MODEL = "deepseek-chat"
# Default timeout: 120s to allow for large file processing and LLM response generation
DEFAULT_TIMEOUT = 120


def get_llm_config():
    """Get LLM configuration from environment variables."""
    return {
        "api_key": os.getenv("LLM_API_KEY", ""),
        "api_url": os.getenv("LLM_API_URL", DEFAULT_API_URL),
        "model": os.getenv("LLM_MODEL", DEFAULT_MODEL),
        "timeout": int(os.getenv("LLM_TIMEOUT", str(DEFAULT_TIMEOUT))),
    }


async def call_llm(messages: list[dict], temperature: float = 0.3) -> str:
    """
    Call the LLM API with given messages and return the response content.

    Args:
        messages: List of message dicts with 'role' and 'content'.
        temperature: Sampling temperature (default 0.3 for more deterministic output).

    Returns:
        The text content of the LLM response.

    Raises:
        ValueError: If API key is not configured.
        httpx.HTTPStatusError: If API returns an error status.
    """
    config = get_llm_config()

    if not config["api_key"]:
        raise ValueError(
            "LLM API key is not configured. "
            "Please set the LLM_API_KEY environment variable."
        )

    headers = {
        "Authorization": f"Bearer {config['api_key']}",
        "Content-Type": "application/json",
    }

    payload = {
        "model": config["model"],
        "messages": messages,
        "temperature": temperature,
    }

    logger.info("Calling LLM API: %s (model: %s)", config["api_url"], config["model"])

    async with httpx.AsyncClient(timeout=config["timeout"]) as client:
        response = await client.post(
            config["api_url"],
            headers=headers,
            json=payload,
        )
        response.raise_for_status()

    data = response.json()
    content = data["choices"][0]["message"]["content"]
    logger.info("LLM response received (%d chars)", len(content))
    return content

import json
import requests
from config import OLLAMA_BASE_URL

def call_ollama_json(model: str, system: str, prompt: str) -> dict:
    """
    Calls Ollama's /api/chat endpoint and returns parsed JSON.
    Raises on connection error or JSON parse failure.
    """
    payload = {
        "model": model,
        "stream": False,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user",   "content": prompt},
        ],
    }
    resp = requests.post(
        f"{OLLAMA_BASE_URL}/api/chat",
        json=payload,
        timeout=120,
    )
    resp.raise_for_status()
    content = resp.json()["message"]["content"]

    # Strip markdown fences if model wraps JSON in ```
    content = content.strip()
    if content.startswith("```"):
        content = content.split("```")[1]
        if content.startswith("json"):
            content = content[4:]
    return json.loads(content.strip())

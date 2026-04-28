#!/usr/bin/env python3
"""Mem0 integration for SlideGen personalization."""

from mem0 import MemoryClient
import json
import os
from dotenv import load_dotenv

load_dotenv()

client = MemoryClient(api_key=os.environ.get("MEM0_API_KEY"))


def get_user_profile(user_id: str) -> dict | None:
    results = client.get_all(filters={"user_id": user_id})
    if not results or results['count'] == 0:
        return None
    return results


def store_user_profile(user_id: str, profile: dict) -> None:
    client.add(json.dumps(profile), user_id=user_id)
    print(f"Stored profile for {user_id} in Mem0")


def get_or_create_profile_path(user_id: str, profile: dict) -> str:
    from pathlib import Path
    tmp_dir = Path("tmp/profiles")
    tmp_dir.mkdir(parents=True, exist_ok=True)
    profile_path = tmp_dir / f"{user_id}_mem0_profile.json"
    reconstructed = {
        "author_id": user_id,
        "source": "mem0",
        "memories": [m['memory'] for m in profile['results']]
    }
    profile_path.write_text(json.dumps(reconstructed, indent=2), encoding="utf-8")
    return str(profile_path)


def update_user_profile(user_id: str, feedback: str) -> None:
    client.add(feedback, user_id=user_id)
    print(f"Updated profile for {user_id} in Mem0")


def delete_user_profile(user_id: str) -> None:
    client.delete_all(filters={"user_id": user_id})
    print(f"Deleted all memories for {user_id}")
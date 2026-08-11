# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

"""Send commands to the bot over Direct Line and print the replies.

Usage::

    python tools/webchat/dl_test.py help channel "agents sdk react"

Reads DIRECTLINE_SECRET from the environment or tools/webchat/.env.
"""

from __future__ import annotations

import json
import os
import sys
import time
import urllib.request
from pathlib import Path

BASE = "https://directline.botframework.com/v3/directline"
_SCHEME = "Bearer "
HERE = Path(__file__).parent


def _load_secret() -> str:
    secret = os.environ.get("DIRECTLINE_SECRET", "")
    env_file = HERE / ".env"
    if not secret and env_file.exists():
        for line in env_file.read_text().splitlines():
            if line.strip().startswith("DIRECTLINE_SECRET="):
                secret = line.split("=", 1)[1].strip()
                break
    if not secret:
        raise SystemExit(f"DIRECTLINE_SECRET is not set (env or {env_file}).")
    return secret


def call(method: str, url: str, token: str, body: dict | None = None) -> dict:
    data = json.dumps(body).encode() if body is not None else b""
    req = urllib.request.Request(
        url,
        method=method,
        data=data,
        headers={"Authorization": _SCHEME + token, "Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=20) as resp:
        raw = resp.read()
    return json.loads(raw) if raw else {}


def describe(activity: dict) -> list[str]:
    lines = []
    if activity.get("text"):
        lines.append(activity["text"])
    for att in activity.get("attachments") or []:
        content = att.get("content")
        texts = []
        if isinstance(content, dict) and isinstance(content.get("body"), list):
            texts = [
                b["text"] for b in content["body"]
                if isinstance(b, dict) and b.get("text")
            ]
        summary = " | ".join(t[:60] for t in texts[:4])
        lines.append(f"[attachment] {att.get('contentType', '?')} {summary}".rstrip())
    if not lines:
        lines.append(f"(empty activity type={activity.get('type')})")
    return lines


def main(messages: list[str]) -> None:
    secret = _load_secret()
    conv = call("POST", f"{BASE}/conversations", secret)
    cid, token = conv["conversationId"], conv["token"]
    print(f"conversation {cid[:20]}…\n")

    watermark = None
    for msg in messages:
        call(
            "POST",
            f"{BASE}/conversations/{cid}/activities",
            token,
            {"type": "message", "from": {"id": "webchat-tester"}, "text": msg},
        )
        time.sleep(3.5)
        url = f"{BASE}/conversations/{cid}/activities"
        if watermark:
            url += f"?watermark={watermark}"
        res = call("GET", url, token)
        watermark = res.get("watermark")

        print(f"→ {msg!r}")
        for activity in res["activities"]:
            if activity["from"]["id"] == "webchat-tester":
                continue
            for line in describe(activity):
                print(f"   ← {line}")
        print()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    main(sys.argv[1:])

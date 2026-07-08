#!/usr/bin/env python3
"""Start CMS AI Quoting and open the browser only after the server is healthy."""
from __future__ import annotations

import subprocess
import sys
import time
import urllib.error
import urllib.request
import webbrowser

HOST = "127.0.0.1"
PORT = 8000
URL = f"http://{HOST}:{PORT}"
HEALTH = f"{URL}/api/health"


def wait_for_health(timeout_sec: float = 45.0) -> bool:
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(HEALTH, timeout=1.5) as resp:
                if resp.status == 200:
                    return True
        except (urllib.error.URLError, TimeoutError, OSError):
            pass
        time.sleep(0.4)
    return False


def main() -> int:
    print(f"Starting CMS AI Quoting on {URL} (local machine only)...")
    proc = subprocess.Popen(
        [
            sys.executable,
            "-m",
            "uvicorn",
            "app.main:app",
            "--host",
            HOST,
            "--port",
            str(PORT),
        ]
    )
    try:
        if wait_for_health():
            print(f"Server ready — opening {URL}")
            try:
                webbrowser.open(URL)
            except Exception:
                pass
        else:
            print("ERROR: Server did not become healthy in time.")
            proc.terminate()
            return 1
        return proc.wait()
    except KeyboardInterrupt:
        proc.terminate()
        return 0


if __name__ == "__main__":
    raise SystemExit(main())

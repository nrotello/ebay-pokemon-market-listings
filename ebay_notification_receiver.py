#!/usr/bin/env python3
"""Small local receiver for eBay's account-deletion endpoint challenge."""
from __future__ import annotations

import hashlib
import json
import os
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse, parse_qs

from dotenv import load_dotenv

load_dotenv()
TOKEN = os.getenv("EBAY_NOTIFICATION_TOKEN", "")
ENDPOINT = os.getenv("EBAY_NOTIFICATION_ENDPOINT", "")
HOST = "127.0.0.1"
PORT = 8787


class Handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        challenge = parse_qs(urlparse(self.path).query).get("challenge_code", [""])[0]
        if not challenge or not TOKEN or not ENDPOINT:
            self.send_error(400, "Missing challenge or endpoint configuration")
            return
        digest = hashlib.sha256(f"{challenge}{TOKEN}{ENDPOINT}".encode("utf-8")).hexdigest()
        body = json.dumps({"challengeResponse": digest}).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self) -> None:
        # This MVP does not retain eBay user data. Acknowledge the event promptly.
        length = int(self.headers.get("Content-Length", "0"))
        if length:
            self.rfile.read(length)
        self.send_response(200)
        self.end_headers()

    def log_message(self, fmt: str, *args: object) -> None:
        print(f"eBay notification receiver: {fmt % args}")


if __name__ == "__main__":
    if not TOKEN or not ENDPOINT:
        raise SystemExit("Set EBAY_NOTIFICATION_TOKEN and EBAY_NOTIFICATION_ENDPOINT in .env first.")
    print(f"Listening on http://{HOST}:{PORT}; public endpoint: {ENDPOINT}")
    HTTPServer((HOST, PORT), Handler).serve_forever()

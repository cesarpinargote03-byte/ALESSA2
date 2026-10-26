from __future__ import annotations

import os

os.environ.setdefault("MICROSOFT_CLIENT_ID", "test-client")
os.environ.setdefault("MICROSOFT_CLIENT_SECRET", "test-client-secret")
os.environ.setdefault("MICROSOFT_ADMIN_EMAIL", "admin@example.com")
os.environ.setdefault(
    "TOKEN_CACHE_ENCRYPTION_KEY",
    "MDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDA=",
)
os.environ.setdefault("SESSION_SECRET_KEY", "test-only-session-secret-key-at-least-32")

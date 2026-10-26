from __future__ import annotations

import os
import tempfile
from contextlib import contextmanager
from pathlib import Path
from typing import Generator

from cryptography.fernet import Fernet, InvalidToken
from filelock import FileLock
from msal import SerializableTokenCache


class TokenCacheError(RuntimeError):
    """Raised when the persisted MSAL cache cannot be read safely."""


class EncryptedTokenCache:
    def __init__(self, path: Path, encryption_key: str) -> None:
        self.path = path
        self._fernet = Fernet(encryption_key.encode("ascii"))
        self._lock_path = path.with_suffix(path.suffix + ".lock")

    @contextmanager
    def session(self) -> Generator[SerializableTokenCache, None, None]:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with FileLock(str(self._lock_path)):
            cache = SerializableTokenCache()
            if self.path.exists():
                try:
                    encrypted = self.path.read_bytes()
                    serialized = self._fernet.decrypt(encrypted).decode("utf-8")
                    cache.deserialize(serialized)
                except (InvalidToken, UnicodeDecodeError, ValueError, TypeError) as exc:
                    raise TokenCacheError(
                        "The encrypted MSAL token cache is invalid or cannot be decrypted."
                    ) from exc

            yield cache

            if cache.has_state_changed:
                self._write_encrypted(cache.serialize().encode("utf-8"))

    def _write_encrypted(self, serialized_cache: bytes) -> None:
        encrypted = self._fernet.encrypt(serialized_cache)
        file_descriptor, temporary_name = tempfile.mkstemp(
            prefix=f".{self.path.name}.",
            dir=self.path.parent,
        )
        temporary_path = Path(temporary_name)
        try:
            with os.fdopen(file_descriptor, "wb") as temporary_file:
                temporary_file.write(encrypted)
                temporary_file.flush()
                os.fsync(temporary_file.fileno())
            if os.name != "nt":
                temporary_path.chmod(0o600)
            os.replace(temporary_path, self.path)
        finally:
            temporary_path.unlink(missing_ok=True)

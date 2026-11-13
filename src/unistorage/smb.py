# pyright: reportMissingTypeStubs=false
# pyright: reportUnknownMemberType=false
# pyright: reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false
# pyright: reportUnknownParameterType=false

import io
import os
from collections.abc import Generator, Iterable
from contextlib import AbstractContextManager, contextmanager
from pathlib import Path
from typing import BinaryIO, Literal, TextIO, overload

import smbclient

from unistorage.abstract import Storage
from unistorage.exceptions import StorageError


class SMBStorage(Storage):
    r"""Dynamic SMB storage using smbclient.
    
    Credentials are reused for any server, extracted from the UNC path.
    Example:
        smb = SMBStorage(username="user", password="pw")
        smb.save(r"\\pr-s000\public\logs.xlsx", b"hello")
        data = smb.read(r"\\another-server\share\file.txt")
    """

    def __init__(self, username: str, password: str) -> None:
        self.username = username
        self.password = password
        self._registered: set[str] = set()

    def _ensure_session(self, path: str | Path) -> None:
        path = str(path)
        if path.startswith("\\\\"):
            server = path.split("\\")[2]
        elif path.startswith("//"):
            server = path.split("/")[2]
        else:
            raise ValueError(f"Invalid SMB path: {path}")

        if server not in self._registered:
            allowed_servers_str = os.getenv("SMB_ALLOWED_SERVERS", "").strip()
            if allowed_servers_str:
                allowed = {s.strip().lower() for s in allowed_servers_str.split(",") if s.strip()}
                if "*" not in allowed and server.lower() not in allowed:
                    raise PermissionError(f"Acesso ao servidor SMB '{server}' não autorizado pela política de segurança.")

            smbclient.register_session(server, username=self.username, password=self.password)
            self._registered.add(server)

    def move(self, src: str, dst: str) -> None:
        """Move or rename a file on an SMB share."""
        try:
            self._ensure_session(src)
            self._ensure_session(dst)
            smbclient.makedirs(os.path.dirname(dst), exist_ok=True)

            # Try rename first (fast path)
            try:
                smbclient.rename(src, dst)
                return
            except Exception:
                pass  # fallback to copy+delete

            # Fallback copy+delete
            with smbclient.open_file(src, mode="rb") as fsrc:
                with smbclient.open_file(dst, mode="wb") as fdst:
                    for chunk in iter(lambda: fsrc.read(1024 * 1024), b""):
                        fdst.write(chunk)
            smbclient.remove(src)
        except Exception as e:
            raise StorageError(f"SMB move failed from {src} to {dst}: {e}") from e

    def write(self, path: str | Path, data: bytes) -> None:
        self._ensure_session(path)
        try:
            smbclient.makedirs(os.path.dirname(path), exist_ok=True)
            with smbclient.open_file(path, mode="wb") as f:
                f.write(data)
        except Exception as e:
            raise StorageError(f"SMB save failed: {e}") from e
        
    def remove(self, path: str | Path) -> None:
        """Remove a file or directory from an SMB share."""
        path = str(path)
        try:
            self._ensure_session(path)
            if not smbclient.smbclient.path.exists(path):
                raise StorageError(f"SMB remove failed: path does not exist ({path})")

            if smbclient.smbclient.path.isdir(path):
                smbclient.rmdir(path)
            else:
                smbclient.remove(path)

        except Exception as e:
            raise StorageError(f"SMB remove failed for {path}: {e}") from e

    def read(self, path: str | Path) -> bytes:
        self._ensure_session(path)
        try:
            with smbclient.open_file(path, mode="rb") as f:
                return f.read()
        except Exception as e:
            raise StorageError(f"SMB read failed: {e}") from e

    def exists(self, path: str | Path) -> bool:
        self._ensure_session(path)
        return smbclient.smbclient.path.exists(path)

    def listdir(self, path: str | Path) -> list[str]:
        self._ensure_session(path)
        try:
            return smbclient.listdir(path)
        except Exception as e:
            raise StorageError(f"SMB listdir failed: {e}") from e

    def makedirs(self, path: str | Path, exist_ok: bool = True) -> None:
        self._ensure_session(path)
        try:
            smbclient.makedirs(path, exist_ok=exist_ok)
        except Exception as e:
            raise StorageError(f"SMB makedirs failed: {e}") from e

    def isdir(self, path: str | Path) -> bool:
        self._ensure_session(path)
        return smbclient.smbclient.path.isdir(path)
    
    def isfile(self, path: str | Path) -> bool:
        self._ensure_session(path)
        return smbclient.smbclient.path.isfile(path)
    
    def walk(self, path: str | Path) -> Iterable[tuple[str, list[str], list[str]]]:
        """Generator yielding (dirpath, dirnames, filenames) for SMB paths."""
        self._ensure_session(path)
        path_str = str(path)
        try:
            dirs: list[str] = []
            files: list[str] = []
            for entry in smbclient.scandir(path_str):
                if entry.is_dir():
                    dirs.append(str(entry.name))
                else:
                    files.append(str(entry.name))
            yield path_str, dirs, files
            for d in dirs:
                new_path = os.path.join(path_str, d)
                yield from self.walk(new_path)
        except Exception as e:
            raise StorageError(f"SMB walk failed for {path}: {e}") from e
        
    def copy(self, src: str, dst: str) -> None:
        """Copy a file from src to dst within SMB storage."""
        self._ensure_session(src)
        self._ensure_session(dst)
        try:
            data = self.read(src)
            self.makedirs(os.path.dirname(dst), exist_ok=True)
            self.write(dst, data)
        except Exception as e:
            raise StorageError(f"SMB copy failed from {src} to {dst}: {e}") from e
    
    @overload
    def open(self, path: str | Path, mode: Literal["rb", "wb", "ab", "r+b", "w+b", "a+b"]) -> AbstractContextManager[BinaryIO]:
        ...
    @overload
    def open(self, path: str | Path, mode: Literal["r", "w", "a", "rt", "wt", "at", "r+", "w+", "a+"] = "r", encoding: str | None = None) -> AbstractContextManager[TextIO]:
        ...
    @overload
    def open(self, path: str | Path, mode: str, encoding: str | None = None) -> AbstractContextManager[BinaryIO | TextIO]:
        ...

    @contextmanager
    def open(self, path: str | Path, mode: str = "r", encoding: str | None = None) -> Generator[BinaryIO | TextIO]:
        """Open an SMB path as a file-like object.

        - If 'b' in mode -> yields a binary file (BinaryIO).
        - Otherwise -> yields a TextIO[str] wrapping a binary stream.
        """
        path = str(path)
        self._ensure_session(path)

        # Ensure parent dirs for write modes
        if any(ch in mode for ch in ("w", "a", "+")):
            try:
                smbclient.makedirs(os.path.dirname(path), exist_ok=True)
            except Exception:
                # best-effort; ignore or raise if you prefer strictness
                pass

        # choose raw mode (binary) for the underlying open_file
        raw_mode = mode if "b" in mode else (mode + "b")  # 'r' -> 'rb', 'w' -> 'wb'
        raw = None
        try:
            raw = smbclient.open_file(path, mode=raw_mode) # type: ignore
            if "b" in mode:
                # binary mode: yield raw binary file
                try:
                    yield raw
                finally:
                    raw.close()
            else:
                # text mode: wrap binary file in TextIOWrapper
                text = io.TextIOWrapper(raw, encoding=encoding or "utf-8")
                try:
                    yield text
                finally:
                    # closing text wrapper closes the underlying raw stream
                    text.close()
        except Exception as e:
            raise StorageError(f"SMB open failed ({mode}) for {path}: {e}") from e
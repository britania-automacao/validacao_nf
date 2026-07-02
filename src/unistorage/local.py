import os
import shutil
from collections.abc import Generator
from contextlib import AbstractContextManager, contextmanager
from pathlib import Path
from typing import BinaryIO, Literal, TextIO, overload

from unistorage.abstract import Storage
from unistorage.exceptions import StorageError


class LocalStorage(Storage):
    """Local filesystem storage."""

    def move(self, src: str, dst: str) -> None:
        """Move or rename a local file or directory."""
        try:
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.move(src, dst)
        except Exception as e:
            raise StorageError(f"Local move failed from {src} to {dst}: {e}") from e

    def write(self, path: str | Path, data: bytes) -> None:
        self.makedirs(os.path.dirname(path), exist_ok=True)
        try:
            tmp_path = str(path) + ".tmp"
            with open(tmp_path, "wb") as f:
                f.write(data)
            os.replace(tmp_path, path)
        except Exception as e:
            raise StorageError(f"Local save failed: {e}") from e
        
    def remove(self, path: str | Path) -> None:
        """Remove a file or directory from local storage."""
        try:
            if not os.path.exists(path):
                raise StorageError(f"Local remove failed: path does not exist ({path})")

            if os.path.isfile(path) or os.path.islink(path):
                os.remove(path)
            elif os.path.isdir(path):
                shutil.rmtree(path)
            else:
                raise StorageError(f"Local remove failed: unsupported path type ({path})")

        except Exception as e:
            raise StorageError(f"Local remove failed for {path}: {e}") from e

    def read(self, path: str | Path) -> bytes:
        try:
            with open(path, "rb") as f:
                return f.read()
        except Exception as e:
            raise StorageError(f"Local read failed: {e}") from e

    def exists(self, path: str | Path) -> bool:
        return os.path.exists(path)

    def listdir(self, path: str | Path) -> list[str]:
        try:
            return os.listdir(path)
        except Exception as e:
            raise StorageError(f"Local listdir failed: {e}") from e

    def makedirs(self, path: str | Path, exist_ok: bool = True) -> None:
        try:
            os.makedirs(path, exist_ok=exist_ok)
        except Exception as e:
            raise StorageError(f"Local makedirs failed: {e}") from e

    def isdir(self, path: str | Path) -> bool:
        return os.path.isdir(path)
    
    def isfile(self, path: str | Path) -> bool:
        return os.path.isfile(path)
    
    def walk(self, path: str | Path) -> Generator[tuple[str, list[str], list[str]]]:
        """Generator yielding (dirpath, dirnames, filenames) like os.walk()."""
        try:
            yield from os.walk(path)
        except Exception as e:
            raise StorageError(f"Local walk failed: {e}") from e

    def copy(self, src: str, dst: str) -> None:
        try:
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            # shutil.copy2 preserves metadata (mtime, permissions)
            shutil.copy2(src, dst)
        except Exception as e:
            raise StorageError(f"Local copy failed from {src} to {dst}: {e}") from e
        
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
        try:
            if "b" in mode:
                f: BinaryIO = open(path, mode)  # type: ignore[assignment]
            else:
                f: TextIO = open(path, mode, encoding=encoding or "utf-8")  # type: ignore[assignment]
            try:
                yield f
            finally:
                f.close()
        except Exception as e:
            raise StorageError(f"Local open failed ({mode}) for {path}: {e}") from e
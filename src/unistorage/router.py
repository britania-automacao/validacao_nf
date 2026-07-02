from collections.abc import Generator, Iterable
from contextlib import AbstractContextManager, contextmanager
from pathlib import Path
from typing import BinaryIO, Literal, TextIO, overload

from unistorage.abstract import Storage
from unistorage.exceptions import StorageError
from unistorage.local import LocalStorage
from unistorage.smb import SMBStorage


class StorageRouter(Storage):
    """
    Routes file operations (open, read, write, etc.) to the correct backend
    based on the path.

    Example:
        router = StorageRouter(
            local=LocalStorage(),
            smb=SMBStorage(username="user", password="pw")
        )

        router.write("data.txt", b"123")                      # → Local
        router.write("//server/share/f.txt", b"abc")          # → SMB
    """

    def __init__(self, local: LocalStorage, smb: SMBStorage | None = None):
        self.local = local
        self.smb = smb

    def _select_storage(self, path: str | Path) -> tuple[Storage, str]:
        """Decide which storage to use based on path."""
        path_str = str(path)
        if path_str.startswith("\\\\") or path_str.startswith("//"):
            if not self.smb:
                raise StorageError(f"SMB path but no SMB storage configured: {path_str}")
            return self.smb, path_str
        return self.local, path_str

    # ---- core operations ----
    def write(self, path: str | Path, data: bytes) -> None:
        storage, real_path = self._select_storage(path)
        return storage.write(real_path, data)

    def read(self, path: str | Path) -> bytes:
        storage, real_path = self._select_storage(path)
        return storage.read(real_path)
    
    def copy(self, src: str | Path, dst: str | Path) -> None:
        """Copy file from src to dst (possibly across storages)."""
        src_storage, src_real = self._select_storage(src)
        dst_storage, dst_real = self._select_storage(dst)

        if src_storage is dst_storage:
            return src_storage.copy(src_real, dst_real)
        
        # Cross-storage copy (streamed)
        with src_storage.open(src_real, "rb") as fsrc:
            with dst_storage.open(dst_real, "wb") as fdst:
                for chunk in iter(lambda: fsrc.read(1024 * 1024), b""):
                    fdst.write(chunk)


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
        storage, real_path = self._select_storage(path)
        with storage.open(real_path, mode=mode, encoding=encoding) as f:
            yield f

    def exists(self, path: str | Path) -> bool:
        storage, real_path = self._select_storage(path)
        return storage.exists(real_path)

    def listdir(self, path: str | Path) -> list[str]:
        storage, real_path = self._select_storage(path)
        return storage.listdir(real_path)

    def makedirs(self, path: str | Path, exist_ok: bool = True) -> None:
        storage, real_path = self._select_storage(path)
        return storage.makedirs(real_path, exist_ok=exist_ok)

    def remove(self, path: str | Path) -> None:
        storage, real_path = self._select_storage(path)
        return storage.remove(real_path)

    def isdir(self, path: str | Path) -> bool:
        storage, real_path = self._select_storage(path)
        return storage.isdir(real_path)

    def move(self, src: str | Path, dst: str | Path) -> None:
        src_storage, src_real = self._select_storage(src)
        dst_storage, dst_real = self._select_storage(dst)

        if src_storage is dst_storage:
            return src_storage.move(src_real, dst_real)

        # cross-storage move (streamed)
        with src_storage.open(src_real, "rb") as fsrc:
            with dst_storage.open(dst_real, "wb") as fdst:
                for chunk in iter(lambda: fsrc.read(1024 * 1024), b""):
                    fdst.write(chunk)
        src_storage.remove(src_real)

    def isfile(self, path: str | Path) -> bool:
        storage, real_path = self._select_storage(path)
        return storage.isfile(real_path)
    
    def walk(self, path: str | Path) -> Iterable[tuple[str, list[str], list[str]]]:
        storage, real_path = self._select_storage(path)
        return storage.walk(real_path)
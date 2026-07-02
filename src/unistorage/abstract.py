from abc import ABC, abstractmethod
from collections.abc import Iterable
from contextlib import AbstractContextManager
from pathlib import Path
from typing import BinaryIO, Literal, TextIO, overload


class Storage(ABC):
    @abstractmethod
    def move(self, src: str, dst: str) -> None: ...

    @abstractmethod
    def write(self, path: str | Path, data: bytes) -> None: ...

    @abstractmethod
    def remove(self, path: str | Path) -> None: ...

    @abstractmethod
    def read(self, path: str | Path) -> bytes: ...

    @abstractmethod
    def exists(self, path: str | Path) -> bool: ...

    @abstractmethod
    def listdir(self, path: str | Path) -> list[str]: ...

    @abstractmethod
    def makedirs(self, path: str | Path, exist_ok: bool = True) -> None: ...

    @abstractmethod
    def isdir(self, path: str | Path) -> bool: ...

    @abstractmethod
    def isfile(self, path: str | Path) -> bool: ...

    @abstractmethod
    def walk(self, path: str | Path) -> Iterable[tuple[str, list[str], list[str]]]: ...

    @abstractmethod
    def copy(self, src: str, dst: str) -> None: ...

    @overload
    def open(self, path: str | Path, mode: Literal["rb", "wb", "ab", "r+b", "w+b", "a+b"]) -> AbstractContextManager[BinaryIO]:
        ...
    @overload
    def open(self, path: str | Path, mode: Literal["r", "w", "a", "rt", "wt", "at", "r+", "w+", "a+"] = "r", encoding: str | None = None) -> AbstractContextManager[TextIO]:
        ...
    @overload
    def open(self, path: str | Path, mode: str, encoding: str | None = None) -> AbstractContextManager[BinaryIO | TextIO]:
        ...

    @abstractmethod
    def open(self, path: str | Path, mode: str = "r", encoding: str | None = None) -> AbstractContextManager[BinaryIO | TextIO]:
        """Open a file-like object on the underlying storage.

        Text modes (e.g. "r", "w", "rt") should return a TextIO[str].
        Binary modes (e.g. "rb", "wb") should return a BinaryIO.
        """
        ...
from __future__ import annotations

import os
import shutil
import tempfile
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator


PRIVATE_FILE_MODE = 0o600
PRIVATE_DIRECTORY_MODE = 0o700


def ensure_private_directory(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True, mode=PRIVATE_DIRECTORY_MODE)
    os.chmod(path, PRIVATE_DIRECTORY_MODE)


def copy_private_file(source: Path, destination: Path) -> None:
    descriptor = os.open(destination, os.O_WRONLY | os.O_CREAT | os.O_EXCL, PRIVATE_FILE_MODE)
    os.close(descriptor)
    try:
        shutil.copyfile(source, destination)
        os.chmod(destination, PRIVATE_FILE_MODE)
    except Exception:
        destination.unlink(missing_ok=True)
        raise


@contextmanager
def private_temporary_path(destination: Path) -> Iterator[Path]:
    descriptor, raw_path = tempfile.mkstemp(
        dir=destination.parent,
        prefix=f".{destination.stem}.",
        suffix=f".tmp{destination.suffix}",
    )
    os.close(descriptor)
    temporary = Path(raw_path)
    os.chmod(temporary, PRIVATE_FILE_MODE)
    try:
        yield temporary
    finally:
        temporary.unlink(missing_ok=True)


def replace_private_file(temporary: Path, destination: Path) -> None:
    os.chmod(temporary, PRIVATE_FILE_MODE)
    os.replace(temporary, destination)
    os.chmod(destination, PRIVATE_FILE_MODE)

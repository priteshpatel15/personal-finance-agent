from __future__ import annotations

from typing import Any, Protocol


class WritableCell(Protocol):
    value: Any
    data_type: str
    hyperlink: Any


def set_literal_cell(cell: WritableCell, value: object) -> None:
    """Write external text as an Excel string, never as an executable formula."""

    cell.hyperlink = None
    cell.value = value
    if isinstance(value, str):
        cell.data_type = "s"

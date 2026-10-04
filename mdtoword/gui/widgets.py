"""Виджеты приёма файлов перетаскиванием: очередь и зона-подсказка."""

from __future__ import annotations

from typing import Any

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QDragEnterEvent, QDragMoveEvent, QDropEvent, QMouseEvent
from PyQt6.QtWidgets import QLabel, QListWidget, QWidget


def dropped_local_paths(event: Any | None) -> list[str]:
    """Local filesystem paths carried by a drop event, if any."""
    if event is None:
        return []
    mime_data = event.mimeData()
    if mime_data is None:
        return []
    return [url.toLocalFile() for url in mime_data.urls() if url.isLocalFile()]


def accept_local_paths_event(event: Any | None) -> None:
    """Accept a drag event that carries at least one local filesystem path."""
    if event is None:
        return
    if dropped_local_paths(event):
        event.acceptProposedAction()
    else:
        event.ignore()


class DropFileList(QListWidget):
    """A queue widget that accepts files and directories from the desktop."""

    paths_dropped = pyqtSignal(list)

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, e: QDragEnterEvent | None) -> None:
        accept_local_paths_event(e)

    def dragMoveEvent(self, e: QDragMoveEvent | None) -> None:
        accept_local_paths_event(e)

    def dropEvent(self, event: QDropEvent | None) -> None:
        if event is None:
            return
        paths = dropped_local_paths(event)
        if paths:
            self.paths_dropped.emit(paths)
            event.acceptProposedAction()
        else:
            event.ignore()


class DropZoneLabel(QLabel):
    """Clickable drop hint that doubles as a file-picker button."""

    clicked = pyqtSignal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("drop-zone")
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setWordWrap(True)
        self.setMinimumHeight(80)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def mouseReleaseEvent(self, event: QMouseEvent | None) -> None:
        if (
            event is not None
            and event.button() == Qt.MouseButton.LeftButton
            and self.rect().contains(event.position().toPoint())
        ):
            self.clicked.emit()
        super().mouseReleaseEvent(event)

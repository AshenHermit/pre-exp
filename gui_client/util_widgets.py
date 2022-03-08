import os
from pathlib import Path
import sys
from PyQt5.QtWidgets import * 
from PyQt5 import QtCore, QtGui
from PyQt5.QtGui import *
from PyQt5.QtCore import * 
from builder import *

from .settings import storage
from .references import *

class LeftElidedLabel(QLabel):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)

    def paintEvent(self, event):
        painter = QPainter(self)
        metrics = QFontMetrics(self.font())
        elided = metrics.elidedText(self.text(), Qt.TextElideMode.ElideLeft, self.width())
        painter.drawText(self.rect(), self.alignment(), elided)

class InputPathWidget(QWidget):
    def __init__(self, path:Path=None, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self._path_ref = PathReference(path)

        self.setObjectName("InputPathWidget")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setStyleSheet('QWidget#InputPathWidget{ background-color: rgba(255, 255, 255, .05); }\n'+\
        'QWidget{ background-color: rgba(255, 255, 255, .0); }')

        self.hbox = QHBoxLayout()
        self.setLayout(self.hbox)

        self.path_label = QLabel()
        self.path_label.setWordWrap(True)
        self.hbox.addWidget(self.path_label)
        self.button = QPushButton("...")
        self.button.pressed.connect(self.select_file)
        self.button.setMaximumWidth(32)
        self.hbox.addWidget(self.button)

        self.path = path

    @property
    def path(self)->Path:
        return self._path_ref.path
    @path.setter
    def path(self, path:Path):
        self._path_ref.path = path
        self.path_label.setText(path.as_posix())
    @property
    def path_ref(self):
        return self._path_ref

    def select_file(self):
        options = QFileDialog.Options()

        if self.path.exists():
            dir = str(self.path.resolve().parent)

        filename, _ = QFileDialog.getOpenFileName(
            self, 
            "Открыть", 
            dir, "Excel (*.xls *.xlsx)", 
            options=options)
        if filename:
            print(filename)
            self.path = Path(filename)
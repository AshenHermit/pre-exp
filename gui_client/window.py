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
from .util_widgets import *
from .builders_widgets import *

CFD = Path(__file__).parent.resolve()

class ClientWindow(QMainWindow):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)

        self.styles_filepath = CFD/"styles/dark/stylesheet.qss"
        self.setup_stylesheet()

        #TODO: разделить элементы на объекты

        self.setWindowTitle("Генератор таблиц")
        if storage.window_size is None: storage.window_size = QSize(1357, 608)
        if storage.window_pos is None: storage.window_pos = QPoint(300, 300)
        self.resize(storage.window_size)
        self.move(storage.window_pos)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.vbox = QVBoxLayout()
        self.central_widget.setLayout(self.vbox)

        self.src_file_frame = QFrame()
        self.vbox.addWidget(self.src_file_frame)
        self.src_file_vbox = QVBoxLayout()
        self.src_file_frame.setLayout(self.src_file_vbox)

        # ввод файла с данными для генерации
        self.src_file_vbox.addWidget(QLabel("Файл с данными для генерации"))
        if storage.input_filepath is None: storage.input_filepath = ""
        self.src_file_input = InputPathWidget(Path(storage.input_filepath))
        self.src_file_vbox.addWidget(self.src_file_input)
        
        self.add_data_sheet_name_input()

        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        self.vbox.addWidget(line)

        self.genertors_classes = [
            PrepaidExpenseGeneratorWidget,
            GraphGeneratorWidget,
            ReportGeneratorWidget,
        ]
        self.genertors_widgets = []
        for qgen_cls in self.genertors_classes:
            qgen = qgen_cls(self.src_file_input.path_ref, self.data_sheet_name_ref)
            self.genertors_widgets.append(qgen)
            self.vbox.addWidget(qgen)

        self.vbox.addWidget(BuildAllWidget(self.genertors_widgets))

        self.vbox.addStretch()

    def setup_stylesheet(self):
        stylesheet = self.styles_filepath.read_text()
        styles_folder_path = (CFD/"styles").as_posix()
        stylesheet = stylesheet.replace("url(:", f"url({styles_folder_path}/")
        qApp.setStyleSheet(stylesheet)

    def add_data_sheet_name_input(self):
        """ ввод названия таблицы с данными для генерации """
        self.sheet_name_frame = QFrame()
        self.sheet_name_box = QHBoxLayout()
        self.sheet_name_frame.setLayout(self.sheet_name_box)
        self.vbox.addWidget(self.sheet_name_frame)
        self.sheet_name_box.addWidget(QLabel("Название таблицы с данными:"))
        if storage.data_sheet_name is None: storage.data_sheet_name = "ВСЕ"
        self.data_sheet_name_input = QLineEdit(storage.data_sheet_name)
        self.data_sheet_name_ref = StringReference()
        self.data_sheet_name_input.textChanged.connect(
            self.data_sheet_name_ref.make_widget_updater(self.data_sheet_name_input))
        self.sheet_name_box.addWidget(self.data_sheet_name_input)

    def closeEvent(self, a0: QtGui.QCloseEvent) -> None:
        storage.input_filepath = self.src_file_input.path.resolve().as_posix()
        storage.data_sheet_name = self.data_sheet_name_input.text()

        storage.window_size = self.size()
        storage.window_pos = self.pos()
        
        return super().closeEvent(a0)

def main():
    app = QApplication(sys.argv)
    window = ClientWindow()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()
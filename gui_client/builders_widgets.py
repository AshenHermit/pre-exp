from functools import partial
import os
from pathlib import Path
import shlex
import sys
import traceback
from PyQt5.QtWidgets import * 
from PyQt5 import QtCore, QtGui
from PyQt5.QtGui import *
from PyQt5.QtCore import * 
from builder import *

from .settings import storage
from .references import *

class GeneratorWidget(QWidget):
    def __init__(self, src_filepath_ref:PathReference, data_sheet_name_ref: StringReference, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self.src_filepath_ref = src_filepath_ref
        self.data_sheet_name_ref = data_sheet_name_ref
        self._filepath_to_open = None

        self.vbox = QVBoxLayout()
        self.setLayout(self.vbox)
        self.vbox.setAlignment(Qt.AlignmentFlag.AlignTop)

        self.setObjectName("GeneratorWidget")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setStyleSheet('QWidget#GeneratorWidget{ background-color: rgba(255, 255, 255, .05); }\n'+\
        'QWidget{ background-color: rgba(255, 255, 255, .0); }\n'+\
        'QLabel#Title{font-size: 12pt;}')

        self.title = QLabel("Генератор")
        self.title.setObjectName("Title")
        self.vbox.addWidget(self.title)

        self.hbox = QGridLayout()
        self.hbox.setColumnStretch(0, 1)
        self.hbox.setColumnStretch(1, 1)
        self.vbox.addLayout(self.hbox)
        self.generate_button = QPushButton("Сгенерировать")
        self.generate_button.pressed.connect(self.generate)
        self.hbox.addWidget(self.generate_button, 0, 0)
        self.open_button = QPushButton("Открыть")
        self.open_button.pressed.connect(self.open_output_file)
        self.hbox.addWidget(self.open_button, 0, 1)

        self.status = QLabel("")
        self.vbox.addWidget(self.status)

        self.filepath_to_open = None

    @property
    def src_filepath(self):
        return self.src_filepath_ref.path
    @property
    def data_sheet_name(self):
        return self.data_sheet_name_ref.string

    @property
    def filepath_to_open(self):
        return self._filepath_to_open
    @filepath_to_open.setter
    def filepath_to_open(self, file:Path):
        self._filepath_to_open = file
        self.open_button.setHidden(self._filepath_to_open is None)

    @property
    def title_text(self):
        return self.title.text()
    @title_text.setter
    def title_text(self, text:str):
        self.title.setText(text)
    
    @property
    def status_text(self):
        return self.status.text()
    @status_text.setter
    def status_text(self, text:str):
        self.status.setText(text)
    
    def open_output_file(self):
        if self.filepath_to_open is not None:
            os.startfile(self.filepath_to_open)
    
    def generate(self):
        self.status_text = "генерирую..."
        self.status_text = "готово"

class SheetGeneratorWidget(GeneratorWidget):
    def make_builder(self):
        return ExcelSheetBuilder()

    @property
    def output_file_prefix(self):
        return ""

    @property
    def output_file(self):
        return self.src_filepath.with_name(self.output_file_prefix + " - " + self.src_filepath.name)

    def generate(self):
        src_file = self.src_filepath
        output_file = self.output_file
        self.status_text = "генерирую..."

        def generate():
            try:
                builder = self.make_builder()
                self.writer = ExcelWorkbookWriter(builder, src_file, output_file, self.data_sheet_name)
                self.writer.run()
                self.status_text = "готово"
                self.filepath_to_open = output_file
            except:
                traceback.print_exc()
                self.status_text = "не получилось"
        QTimer.singleShot(100, generate)

class GraphGeneratorWidget(SheetGeneratorWidget):
    def __init__(self, src_filepath_ref: PathReference, data_sheet_name_ref: StringReference, *args, **kwargs) -> None:
        super().__init__(src_filepath_ref, data_sheet_name_ref, *args, **kwargs)
        self.title_text = "График"

    @property
    def output_file_prefix(self):
        return "График"
    
    def make_builder(self):
        return GraphBuilder()

class PrepaidExpenseGeneratorWidget(SheetGeneratorWidget):
    def __init__(self, src_filepath_ref: PathReference, data_sheet_name_ref: StringReference, *args, **kwargs) -> None:
        super().__init__(src_filepath_ref, data_sheet_name_ref, *args, **kwargs)
        self.title_text = "Табель"

    @property
    def output_file_prefix(self):
        return "Табель"
    
    def make_builder(self):
        return PrepaidExpenseBuilder()

class ReportGeneratorWidget(SheetGeneratorWidget):
    def __init__(self, src_filepath_ref: PathReference, data_sheet_name_ref: StringReference, *args, **kwargs) -> None:
        super().__init__(src_filepath_ref, data_sheet_name_ref, *args, **kwargs)
        self.title_text = "Дневной отчет"

        self.day_hbox = QHBoxLayout()
        self.vbox.insertLayout(1, self.day_hbox)
        self.day_hbox.addWidget(QLabel("День:"))
        if storage.report_selected_day is None: storage.report_selected_day = "1"
        self.day_input = QLineEdit(storage.report_selected_day)
        self.day_input.textChanged.connect(self.day_updated)
        self.day_hbox.addWidget(self.day_input)

    def day_updated(self):
        storage.report_selected_day = self.day_input.text()
    
    @property
    def days_range(self):
        got_values = shlex.split(self.day_input.text().replace("-", " "))
        if len(got_values)>0:
            days_range = [int(got_values[min(i, len(got_values)-1)]) for i in range(2)]
        else:
            days_range = [1, 1]
        return days_range
    
    @property
    def output_file_prefix(self):
        return "Дневной отчет"

    def generate(self):
        src_file = self.src_filepath
        self.status_text = "генерирую..."
        
        def generate():
            success_count = 0
            fail_count = 0
            writer = ExcelWorkbookWriter(builder, src_file, None, self.data_sheet_name)
            writer.read_data_file()

            d = self.days_range[0]
            def gen_next():
                nonlocal fail_count, d, success_count
                if d >= self.days_range[1]+1: return

                try:
                    output_file = (self.src_filepath.parent / "Дневные отчеты") / f"{d} - {self.src_filepath.name}"
                    output_file.parent.mkdir(parents=True, exist_ok=True)
                    builder = DayReportBuilder(selected_day=d)
                    writer.generate_to_file(builder, output_file)
                    success_count += 1
                    self.filepath_to_open = output_file.parent
                except:
                    traceback.print_exc()
                    fail_count+=1

                status = f"готово {success_count} файлов"
                if fail_count>0:
                    status += f", пропущено {fail_count}"

                self.status_text = status

                d += 1
                QTimer.singleShot(1, gen_next)
            QTimer.singleShot(1, gen_next)
                
        QTimer.singleShot(100, generate)

class BuildAllWidget(QWidget):
    def __init__(self, generators_widgets:list, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)

        self.generators_widgets = generators_widgets

        self.vbox = QVBoxLayout()
        self.setLayout(self.vbox)
        self.vbox.setAlignment(Qt.AlignmentFlag.AlignTop)

        self.setObjectName("GeneratorWidget")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setStyleSheet('QWidget#GeneratorWidget{ background-color: rgba(255, 255, 255, .05); }\n'+\
        'QWidget{ background-color: rgba(255, 255, 255, .0); }\n'+\
        'QLabel#Title{font-size: 12pt;}')

        self.gen_button = QPushButton("Сгенерировать все")
        self.gen_button.pressed.connect(self.generate_all)
        self.vbox.addWidget(self.gen_button)

    def generate_all(self):
        for gen in reversed(self.generators_widgets):
            gen.generate()
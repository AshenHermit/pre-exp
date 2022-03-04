from builder import GraphBuilder, PrepaidExpenseBuilder
from old_builder import Builder
import sys
import eel
import subprocess
import tkinter as tk
from tkinter import filedialog
import os

from pathlib import Path

rus_locale = {
    "graph": "график",
    "table": "табель"
}

class Client():

    def __init__(self):
        super().__init__()
        self.setup_gui()

    def setup_gui(self):
        eel.init('gui')
        self.setup_api()
        eel.start('index.html')

    def setup_api(self):
        @eel.expose
        def generate(type, input_path):
            input_path = Path(input_path)
            output_path = input_path.with_name(rus_locale[type] + "_" + input_path.with_suffix("").name + ".xlsx")
            print(input_path)
            print(output_path)
            if type == "table":
                builder = Builder(generation_type="table")
                builder.build(input_path, output_path)
                eel.on_generated(type, output_path)
                pass
            elif type == "graph":
                builder = Builder(generation_type="graph")
                builder.build(input_path, output_path)
                eel.on_generated(type, output_path)
                pass

        eel.expose
        def generate(type, input_path):
            input_path = Path(input_path)
            output_path = input_path.with_name(rus_locale[type] + "_" + input_path.with_suffix("").name + ".xlsx")
            print(input_path)
            print(output_path)
            if type == "table":
                builder = GraphBuilder()
                builder.build(input_path, output_path)
                eel.on_generated(type, output_path)
                pass
            elif type == "graph":
                builder = Builder(generation_type="graph")
                builder.build(input_path, output_path)
                eel.on_generated(type, output_path)
                pass

        @eel.expose
        def select_file():
            root = tk.Tk()
            root.withdraw()
            root.wm_attributes('-topmost', 1)
            input_path = filedialog.askopenfilename()
            input_path = input_path.replace("\\", "/")
            return input_path

        @eel.expose
        def open_file_with_default_program(filepath):
            os.startfile('"'+filepath+'"')
        

def main():
    app = Client()

if __name__ == "__main__":
    main()
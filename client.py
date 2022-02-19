import os
from builder import Builder
import sys
import eel
import subprocess

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
            input_path = input_path.replace("\\", "/")
            output_path = input_path[:input_path.rfind("/")] + "/"+type+"_"+ input_path[input_path.rfind("/")+1:input_path.rfind(".")] +".xlsx"
            print(input_path)
            print(output_path)
            if type == "table":
                builder = Builder(generation_type="table")
                # builder.build(input_path, output_path)
                eel.on_generated(type, output_path)
                pass
            elif type == "graph":
                builder = Builder(generation_type="graph")
                # builder.build(input_path, output_path)
                eel.on_generated(type, output_path)
                pass

        @eel.expose
        def open_file_with_default_program(filepath):
            os.startfile(filepath)

def main():
    app = Client()

if __name__ == "__main__":
    main()
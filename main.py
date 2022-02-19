from client import Client
from builder import Builder

import argparse
parser = argparse.ArgumentParser(description='Помощь:')
parser.add_argument('--graph', action="store_true", help='сгенерировать график')
args = parser.parse_args()


import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

def main_console():
    generation_type = ("graph" if args.graph else "table")
    print("тип сборки: " + generation_type)

    builder = Builder(generation_type)
    print("выберите таблицу суткок:")
    input_path = filedialog.askopenfilename()
    if input_path != '':
        print(input_path)
        print()
        # try:
        
        output_path = input_path[:input_path.rfind("/")] + "/"+generation_type+"_"+ input_path[input_path.rfind("/")+1:input_path.rfind(".")] +".xlsx"
        builder.build(input_path, output_path)
        print()
        print("сохранено в файл: {}".format(output_path))

        # except:
        #     print("что-то пошло не так.")

def main_client():
    client = Client()
    client.run()

def main_test():
    builder = Builder()
    builder.translate_to_new_input_table("./input.xlsx", "./декабрь_сутки — копия.xlsx")

if __name__ == "__main__":
    main_console()
    input("нажмите Enter чтобы выйти.")
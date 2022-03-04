from pathlib import Path
import traceback
from builder.builder import ExcelWorkbookWriter
from builder.graph_builder import GraphBuilder
from builder.prepaid_exp_builder import PrepaidExpenseBuilder
from client import Client
from old_builder import Builder

from old_builder import Builder

import argparse

import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

gen_type_locale = {
    "graph": "график",
    "table": "табель"
}

builders = {
    "graph": GraphBuilder(),
    "table": PrepaidExpenseBuilder()
}

def main_console():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', type=str, help='входной xlsx файл', default="")
    parser.add_argument('--type', type=str, help='тип генерации: graph / table')
    parser.add_argument('--sheet-name', type=str, help='название таблицы для сканирования', default="ВСЕ")
    args = parser.parse_args()

    generation_type = args.type
    if generation_type not in builders:
        print(f"неверный тип сборки: \"{generation_type}\"")
        return

    print(f"тип сборки: \"{gen_type_locale[generation_type]}\"")

    try:
        if args.input!="":
            input_path = args.input
        else:
            input_path = filedialog.askopenfilename()
        input_path = Path(input_path).resolve()
        if input_path.exists():
            gen_type_loc = gen_type_locale[generation_type]
            output_path = input_path.with_name(gen_type_loc + " - " + input_path.with_suffix("").name + ".xlsx")
            
            builder = builders[generation_type]
            writer = ExcelWorkbookWriter(builder, input_path, output_path, args.sheet_name)
            writer.run()
    except:
        print("что-то пошло не так")
        traceback.print_exc()

def main_client():
    client = Client()

def main_test():
    builder = Builder()
    builder.translate_to_new_input_table("./input.xlsx", "./декабрь_сутки — копия.xlsx")

if __name__ == "__main__":
    main_console()
    input("нажмите любую кнопку чтобы выйти.")
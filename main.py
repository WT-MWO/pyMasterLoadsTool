# main workflow

# Import loads:
# clean the spreadsheet - done
# Read the loads - importer class - done
# Write the excel - writer class - done
# TODO: check if connection with Robot exist, raise error
# TODO: implement load writing to Robot
# TODO: check if angles of load are imported properly


# Export loads:
# clean all robot loads in the model
# Read the excel - reader class
# export loads -- exporter class

import clr
import time
import tkinter as tk
from tkinter import filedialog
from tkinter.filedialog import askopenfile
import os

clr.AddReference(r"C:\Program Files\Autodesk\Robot Structural Analysis Professional 2023\Exe\Interop.RobotOM.dll")
from RobotOM import *
import RobotOM as rbt

from pymasterloadstool import importer

root = tk.Tk()
root.title("pyMasterLoadsTool")
root.geometry("500x420")
# Main frame
main_frame = tk.Frame(root)
main_frame.pack()


# TODO: Implement GUI
# TODO: Implement combination import/exoport
# TODO: Implement contour loads

# app = RobotApplication()

# path = r"C:\Users\mwo\OneDrive - WoodThilsted Partners\Professional\5_PYTHON\pyMasterLoadsTool\pyMaster_loads_tool.xlsx"

# trigger = 1  # 0  for export, 1 for import

# start_time = time.time()

# if trigger == 1:
#     # import loads
#     import_load = importer.Importer(app, path=path)
#     records = import_load.import_loads()

# else:
#     # read loads
#     read_xls_loads = xlsreader.XlsReader(path)
#     loads = read_xls_loads.read_data()
#     for l in loads:
#         print(l)


def open_file():
    file = filedialog.askopenfile(mode="r", filetypes=[("Excel", "*.xlsx")])
    if file:
        global filepath
        filepath = os.path.abspath(file.name)


# Executing script
def run_script():
    print(filepath)
    pass


def close_window():
    root.destroy()


# print("--- %s seconds ---" % (time.time() - start_time))
# Input frame
input_frame = tk.LabelFrame(main_frame, text="Path input")
input_frame.grid(row=0, column=0, sticky=tk.EW, padx=5, pady=5)


browse_button = tk.Button(input_frame, text="Browse for file", command=open_file, width=60)
browse_button.grid(row=0, column=0, columnspan=2, sticky=tk.EW, padx=5, pady=5)
# browse_button.place(anchor=tk.CENTER, relx=0.5, rely=0.5)

# Labels frame
labels_frame = tk.LabelFrame(main_frame, text="Choose macro")
labels_frame.grid(row=1, column=0, sticky=tk.EW, padx=5, pady=5)
label_1 = tk.Label(labels_frame, text="Import loads and/or combinations")
label_1.grid(row=0, column=0, padx=5, pady=5)
label_1 = tk.Label(labels_frame, text="Export loads and/or combinations")
label_1.grid(row=1, column=0, padx=5, pady=5)
run_button = tk.Button(labels_frame, text="Run", command=run_script, width=30)
run_button.grid(row=0, column=1, sticky=tk.E, padx=5, pady=5)
run_button = tk.Button(labels_frame, text="Run", command=run_script, width=30)
run_button.grid(row=1, column=1, sticky=tk.E, padx=5, pady=5)
# cancel_button = tk.Button(controls_frame, text="Cancel", command=close_window, width=10)
# cancel_button.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)

root.mainloop()

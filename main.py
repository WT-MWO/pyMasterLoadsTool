import clr
import time
import tkinter as tk
from tkinter import filedialog as fd
from tkinter import ttk
from tkinter import messagebox
from tkinter.filedialog import askopenfile
import os

clr.AddReference(r"C:\Program Files\Autodesk\Robot Structural Analysis Professional 2023\Exe\Interop.RobotOM.dll")
from RobotOM import *
import RobotOM as rbt

from pymasterloadstool import importer, exporter


# TODO: Implement combination export
# TODO: Implement contour loads


# --------------------
# Functions
# --------------------


def check_path():
    if path_var.get() == path_prompt:
        messagebox.showwarning(title="Warning", message=path_prompt)


def import_loads():
    check_path()
    status_msg.set("Processing...")
    start_time = time.time()
    # import loads
    import_load = importer.Importer(app, path=path_var.get())
    import_load.import_loads_and_comb()
    end_time = time.time() - start_time
    status_msg.set(
        "Done. Execution time %f" % end_time,
    )


def export_loads():
    status_msg.set("Processing...")
    start_time = time.time()
    check_path()
    print(path_var.get())
    export_loads = exporter.Exporter(app=app, path=path_var.get())
    export_loads._del_all_cases()
    export_loads._del_all_combinations()
    export_loads.export_load_and_cases()
    export_loads.export_combinations()
    end_time = time.time() - start_time
    status_msg.set(
        "Done. Execution time %f" % end_time,
    )


def open_file():
    file = fd.askopenfilename(filetypes=[("Excel", "*.xlsx")])
    path_var.set(file)


def close_window():  # not used
    root.destroy()


# --------------------
# Main GUI code
# --------------------

root = tk.Tk()
root.title("pyMasterLoadsTool")
root.geometry("700x250")
# Main frame
main_frame = tk.Frame(root)
main_frame.pack()

# Initialize Robot connection
app = RobotApplication()


# Global variables
path_prompt = "Please choose file path."
path_var = tk.StringVar(value=path_prompt)
status_msg = tk.StringVar(value=" ")


# Input frame
input_frame = tk.LabelFrame(main_frame, text="Path input")
input_frame.grid(row=0, column=0, sticky=tk.EW, padx=5, pady=5)
path_label = tk.Label(input_frame, textvariable=path_var)
browse_button = tk.Button(input_frame, text="Browse for file", command=open_file, width=60)
# browse_button.grid(row=0, column=0, columnspan=2, sticky=tk.EW, padx=5, pady=5)
browse_button.pack()
path_label.pack()

# Labels frame
labels_frame = tk.LabelFrame(main_frame, text="Choose macro")
labels_frame.grid(row=1, column=0, sticky=tk.EW, padx=5, pady=5)

label_1 = tk.Label(labels_frame, text="Import loads and/or combinations")
label_1.grid(row=0, column=0, padx=5, pady=5)

label_2 = tk.Label(labels_frame, text="Export loads and/or combinations")
label_2.grid(row=1, column=0, padx=5, pady=5)

run_import_button = tk.Button(labels_frame, text="Run", command=import_loads, width=30)
run_import_button.grid(row=0, column=1, sticky=tk.E, padx=5, pady=5)

run_export_button = tk.Button(labels_frame, text="Run", command=export_loads, width=30)

run_export_button.grid(row=1, column=1, sticky=tk.E, padx=5, pady=5)

# label_1.pack(side="left")
# run_import_button.pack(side="left")
# label_2.pack()
# run_export_button.pack()

# Timestamp frame
timestamp_frame = tk.LabelFrame(main_frame, text="Status")
timestamp_frame.grid(row=2, column=0, sticky=tk.EW, padx=5, pady=5)
info_label = tk.Label(
    timestamp_frame,
    textvariable=status_msg,
    padx=5,
    pady=5,
    justify="left",
)
info_label.pack()
# cancel_button = tk.Button(controls_frame, text="Cancel", command=close_window, width=10)
# cancel_button.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)

root.mainloop()

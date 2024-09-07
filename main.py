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
    import_load.import_loads_and_comb(
        import_loads=trigger_import_loads.get(), import_comb=trigger_import_combinations.get()
    )
    end_time = round(time.time() - start_time)
    status_msg.set(
        "Import done. Execution time %f sec." % end_time,
    )


def export_loads():
    status_msg.set("Processing...")
    start_time = time.time()
    check_path()
    print(path_var.get())
    export_loads = exporter.Exporter(app=app, path=path_var.get())
    export_loads._export_load_cases_combinations(trigger_export_loads.get(), trigger_export_combinations.get())
    end_time = round(time.time() - start_time, 0)
    status_msg.set(
        "Export done. Execution time %f sec." % end_time,
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
root.geometry("700x350")
# root.resizable(width=0, height=100)
# Main frame
main_frame = tk.Frame(root)
main_frame.pack()

# Initialize Robot connection
app = RobotApplication()


# Global variables
path_prompt = "Please choose file path."
path_var = tk.StringVar(value=path_prompt)
status_msg = tk.StringVar(value=" ")
trigger_import_loads = tk.IntVar(value=1)
trigger_import_combinations = tk.IntVar(value=1)
trigger_export_loads = tk.IntVar(value=1)
trigger_export_combinations = tk.IntVar(value=1)


# Input frame
input_frame = tk.LabelFrame(main_frame, text="Path input")
input_frame.grid(row=0, column=0, sticky=tk.EW, padx=5, pady=5)
path_label = tk.Label(input_frame, textvariable=path_var, wraplength=600)
browse_button = tk.Button(input_frame, text="Browse for file", command=open_file, width=60)
# browse_button.grid(row=0, column=0, columnspan=2, sticky=tk.EW, padx=5, pady=5)
browse_button.pack()
path_label.pack()

# Labels frame
labels_frame = tk.LabelFrame(main_frame, text="Choose macro")
labels_frame.grid(row=1, column=0, sticky=tk.EW, padx=5, pady=5)

label_1 = tk.Label(labels_frame, text="Import loads:")
label_1.grid(row=0, column=0, padx=5, pady=5)
label_1_1 = tk.Label(labels_frame, text="combinations:")
label_1_1.grid(row=0, column=2, padx=5, pady=5)


# Import, row 0
label_2 = tk.Label(labels_frame, text="Export loads:")
label_2.grid(row=1, column=0, padx=5, pady=5)
label_2 = tk.Label(labels_frame, text="combinations:")
label_2.grid(row=1, column=2, padx=5, pady=5)

checkbox_import_loads = tk.Checkbutton(labels_frame, variable=trigger_import_loads, onvalue=1, offvalue=0)
checkbox_import_loads.grid(row=0, column=1, sticky=tk.E, padx=5, pady=5)
checkbox_import_combinations = tk.Checkbutton(labels_frame, variable=trigger_import_combinations, onvalue=1, offvalue=0)
checkbox_import_combinations.grid(row=0, column=3, sticky=tk.E, padx=5, pady=5)
run_import_button = tk.Button(labels_frame, text="Run", command=import_loads, width=30)
run_import_button.grid(row=0, column=4, sticky=tk.E, padx=5, pady=5)

# Export row 1
checkbox_export_loads = tk.Checkbutton(labels_frame, variable=trigger_export_loads, onvalue=1, offvalue=0)
checkbox_export_loads.grid(row=1, column=1, sticky=tk.E, padx=5, pady=5)
checkbox_export_combinations = tk.Checkbutton(labels_frame, variable=trigger_export_combinations, onvalue=1, offvalue=0)
checkbox_export_combinations.grid(row=1, column=3, sticky=tk.E, padx=5, pady=5)
run_export_button = tk.Button(labels_frame, text="Run", command=export_loads, width=30)
run_export_button.grid(row=1, column=4, sticky=tk.E, padx=5, pady=5)


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


root.mainloop()

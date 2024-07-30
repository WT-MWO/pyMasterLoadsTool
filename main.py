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

clr.AddReference(r"C:\Program Files\Autodesk\Robot Structural Analysis Professional 2023\Exe\Interop.RobotOM.dll")
from RobotOM import *
import RobotOM as rbt

from pymasterloadstool import importer, writer
from pymasterloadstool import xlsreader

app = RobotApplication()

path = r"C:\Users\mwo\OneDrive - WoodThilsted Partners\Professional\5_PYTHON\pyMasterLoadsTool\pyMaster_loads_tool.xlsx"

trigger = 0  # 0  for export, 1 for import

start_time = time.time()

if trigger == 1:
    # import loads
    import_load = importer.Importer(app)
    records = import_load.get_load_records()

    # write loads
    write_loads = writer.Writer(path)
    write_loads.write_data(records)
else:
    # read loads
    read_xls_loads = xlsreader.XlsReader(path)
    loads = read_xls_loads.read_data()
    for l in loads:
        print(l)

print("--- %s seconds ---" % (time.time() - start_time))

# main workflow
# Import loads
# clean the spreadsheet
# Read the loads - importer class
# Write the excel - writer class

# Export loads
# clean all robot loads in the model
# Read the excel - reader class
# export loads -- exporter class

import clr

clr.AddReference(r"C:\Program Files\Autodesk\Robot Structural Analysis Professional 2023\Exe\Interop.RobotOM.dll")
from RobotOM import *
import RobotOM as rbt

from pymasterloadstool import importer, writer

app = RobotApplication()

path = r"C:\Users\mwo\OneDrive - WoodThilsted Partners\Professional\5_PYTHON\pyMasterLoadsTool\pyMaster_loads_tool.xlsx"

# import loads
import_load = importer.Importer(app)
records = import_load.get_load_records()

print(records)

# write loads
write_loads = writer.Writer(path)
write_loads.write_data(records)

# TODO: case number

clr.AddReference(r"C:\Program Files\Autodesk\Robot Structural Analysis Professional 2023\Exe\Interop.RobotOM.dll")
from RobotOM import *
import RobotOM as rbt

from structure import Structure


class Exporter:
    """Exports the loads from the excel to the Robot model"""

    def __init__(self, app):
        super().__init__(app)

import clr
import sys
import math
from .pyLoad import pyLoad, missing_msg

clr.AddReference(r"C:\Program Files\Autodesk\Robot Structural Analysis Professional 2023\Exe\Interop.RobotOM.dll")
from RobotOM import *
import RobotOM as rbt

from .structure import Structure, supported_load_types

U = 1000  # divider to get kN
R = 2  # rounding
rad_to_deg = 180 / math.pi


class Importer(Structure):
    """Reads the loads and combinations from the model, imports them as a list of pyLoad objects."""

    def __init__(self, app):
        super().__init__(app)
        self.supported_load_types = supported_load_types

    # rect = rbt.IRobotLoadRecordType

    def read_cosystem(self, cosys):
        if cosys == 0:
            return "global"
        else:
            return "local"

    def read_calcnode(self, cnode):
        if cnode == 0:
            return "no"
        else:
            return "generated"

    def read_relabs(self, rela):
        if rela == 0:
            return "absolute"
        else:
            return "relative"

    def _check_type(self, rec):
        rec_type = int(rec.Type)
        if rec_type in self.supported_load_types:
            return True
        else:
            return False

    def read_objects(self, rec):
        """Reads the objects to which the load is applied to. For example bar numbers."""
        objects = []
        rect = int(rec.Type)
        # print(rect)
        if rect != 89:  # if not a Body force
            count = rec.Objects.Count
            if count == 0:
                if rect == 69:  # in edge load case objects are empty but do contain text
                    obj = rec.Objects
                    objects.append(obj.ToText())
                else:
                    objects.append(missing_msg)
            else:
                for r in range(1, count + 1):
                    objects.append(rec.Objects.Get(r))
        else:
            if rec.GetValue(12) == 0:
                # TODO: can this read real objects list? it is displayed in robot...seems like this is not possible
                objects.append("all")
            else:
                objects.append("objects")
        return objects

    def _store_load(self, lcase, rec):
        """Reads the load from the model, stores it in pyLoad object and returns them in list"""
        load = pyLoad()
        rec_type = int(rec.Type)  # cast it as a int
        load.LCName = lcase.Name
        load.LCNumber = lcase.Number
        load.type = rec_type
        load.objects = self.read_objects(rec)  # read the objects load is assigned to
        if rec_type == 0:  # nodal force
            load.Name = self.supported_load_types[rec_type]
            load.FX = round(rec.GetValue(0) / U, R)  # I_NFIPRV_FX
            load.FY = round(rec.GetValue(1) / U, R)  # I_NFIPRV_FY
            load.FZ = round(rec.GetValue(2) / U, R)  # I_NFIPRV_FZ
            load.Mx = round(rec.GetValue(3) / U, R)  # I_NFIPRV_MX
            load.My = round(rec.GetValue(4) / U, R)  # I_NFIPRV_MY
            load.Mz = round(rec.GetValue(5) / U, R)  # I_NFIPRV_MZ
            load.alfa = rec.GetValue(8)  # I_NFIPRV_ALPHA
            load.beta = rec.GetValue(9)  # I_NFIPRV_BETA
            load.gamma = rec.GetValue(10)  # I_NFIPRV_GAMMA
            # load.calcnode = self.read_calcnode(rec.GetValue(12))
        elif rec_type == 5:  # uniform load
            load.Name = self.supported_load_types[rec_type]
            load.FX = round(rec.GetValue(0) / U, R)
            load.FY = round(rec.GetValue(1) / U, R)
            load.FZ = round(rec.GetValue(2) / U, R)
            load.Mx = round(rec.GetValue(3) / U, R)
            load.My = round(rec.GetValue(4) / U, R)
            load.Mz = round(rec.GetValue(5) / U, R)
            load.alfa = rec.GetValue(8)
            load.beta = rec.GetValue(9)
            load.gamma = rec.GetValue(10)
            load.cosystem = self.read_cosystem(rec.GetValue(11))
            load.projected = rec.GetValue(12)
            load.absrel = self.read_relabs(rec.GetValue(13))
            load.disY = rec.GetValue(21)
            load.disZ = rec.GetValue(22)
        elif rec_type == 26:  # uniform load on a FE element
            load.Name = self.supported_load_types[rec_type]
            load.PX = round(rec.GetValue(0) / U, R)  # I_URV_PX
            load.PY = round(rec.GetValue(1) / U, R)  # I_URV_PY
            load.PZ = round(rec.GetValue(2) / U, R)  # I_URV_PZ
            load.cosystem = self.read_cosystem(rec.GetValue(11))  # I_URV_LOCAL_SYSTEM
            load.projected = rec.GetValue(12)  # I_URV_PROJECTION
        elif rec_type == 7:  # self-weight
            load.Name = self.supported_load_types[rec_type]
            load.entirestruc = rec.GetValue(15)  # I_BDRV_ENTIRE_STRUCTURE
        elif rec_type == 3:  # point load on a bar
            load.Name = self.supported_load_types[rec_type]
            load.FX = round(rec.GetValue(0) / U, R)  # I_BFCRV_FX
            load.FY = round(rec.GetValue(1) / U, R)  # I_BFCRV_FY
            load.FZ = round(rec.GetValue(2) / U, R)  # I_BFCRV_FZ
            load.Mx = round(rec.GetValue(3) / U, R)  # I_BFCRV_CX
            load.My = round(rec.GetValue(4) / U, R)  # I_BFCRV_CY
            load.Mz = round(rec.GetValue(5) / U, R)  # I_BFCRV_CZ
            load.alfa = rec.GetValue(8)  # I_BFCRV_ALPHA
            load.beta = rec.GetValue(9)  # I_BFCRV_BETA
            load.gamma = rec.GetValue(10)  # I_BFCRV_GAMMA
            load.cosystem = self.read_cosystem(rec.GetValue(11))  # I_BFCRV_LOC
            load.calcnode = self.read_calcnode(rec.GetValue(12))  # I_BFCRV_GENERATE_CALC_NODE
            load.absrel = self.read_relabs(rec.GetValue(13))  # I_BFCRV_REL
            load.disX = rec.GetValue(6)  # I_BFCRV_X
            load.disY = rec.GetValue(21)  # I_BFCRV_OFFSET_Y
            load.disZ = rec.GetValue(22)  # I_BFCRV_OFFSET_Z
        elif rec_type == 6:  # "trapezoidal load (2p)"
            load.Name = self.supported_load_types[rec_type]
            load.PX = round(rec.GetValue(3) / U, R)  # I_BTRV_PX1
            load.PY = round(rec.GetValue(4) / U, R)  # I_BTRV_PY1
            load.PZ = round(rec.GetValue(5) / U, R)  # I_BTRV_PZ1
            load.PX2 = round(rec.GetValue(0) / U, R)  # I_BTRV_PX2
            load.PY2 = round(rec.GetValue(1) / U, R)  # I_BTRV_PY2
            load.PZ2 = round(rec.GetValue(2) / U, R)  # I_BTRV_PZ2
            load.disX = rec.GetValue(7)  # I_BTRV_X1
            load.disX2 = rec.GetValue(6)  # I_BTRV_X2
            load.alfa = rec.GetValue(8)  # I_BTRV_ALPHA
            load.beta = rec.GetValue(9)  # I_BTRV_BETA
            load.gamma = rec.GetValue(10)  # I_BTRV_GAMMA
            load.projected = rec.GetValue(12)  # I_BTRV_PROJECTION
            load.absrel = self.read_relabs(rec.GetValue(13))  # I_BTRV_RELATIVE
        elif rec_type == 69:  # (FE) 2 load on edges
            load.Name = self.supported_load_types[rec_type]
            load.PX = round(rec.GetValue(0) / U, R)  # I_LOERV_PX
            load.PY = round(rec.GetValue(1) / U, R)  # I_LOERV_PY
            load.PZ = round(rec.GetValue(2) / U, R)  # I_LOERV_PZ
            load.Mx = round(rec.GetValue(3) / U, R)  # I_LOERV_MX
            load.My = round(rec.GetValue(4) / U, R)  # I_LOERV_MY
            load.Mz = round(rec.GetValue(5) / U, R)  # I_LOERV_MZ
            load.gamma = rec.GetValue(6)  # I_LOERV_GAMMA
            load.cosystem = self.read_cosystem(rec.GetValue(11))  # I_LOERV_LOCAL_SYSTEM
        elif rec_type == 89:  # "Body forces"
            load.Name = self.supported_load_types[rec_type]
            load.FX = round(rec.GetValue(0), R)
            load.FY = round(rec.GetValue(1), R)
            load.FZ = round(rec.GetValue(2), R)
            load.absrel = self.read_relabs(rec.GetValue(13))  # I_BURV_RELATIVE
        return load

    def get_load_records(self):
        "Returns a list of load records of pyLoad object."
        cases = self.structure.Cases.GetAll()
        records = []
        for i in range(1, cases.Count + 1):
            lcase = rbt.IRobotCase(cases.Get(i))
            # Make sure case is not a combination or not supported type
            if lcase.AnalizeType in self.supported_analize_type:
                # Recast case object as simple
                lcase = rbt.IRobotSimpleCase(cases.Get(i))
                # Loop through the load records in case
                # print(lcase.Records.Count)
                for r in range(1, lcase.Records.Count + 1):
                    # Get load record
                    rec = lcase.Records.Get(r)
                    # Check if load record is not auto generated from claddings loads
                    rec_com = rbt.IRobotLoadRecordCommon(lcase.Records.Get(r))
                    if not rec_com.IsAutoGenerated:
                        # check if record is supported
                        if self._check_type(rec):
                            records.append(self._store_load(lcase, rec))
        return records


if __name__ == "__main__":
    app = RobotApplication()
    # pyrobot = pyARSAReporting(app)
    importer = Importer(app)
    records = importer.get_load_records()
    print(records)
    # for r in records:
    #     print(r)

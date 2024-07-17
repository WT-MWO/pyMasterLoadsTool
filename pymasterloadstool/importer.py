import clr
import sys
from pyLoad import pyLoad

clr.AddReference(r"C:\Program Files\Autodesk\Robot Structural Analysis Professional 2023\Exe\Interop.RobotOM.dll")
from RobotOM import *
import RobotOM as rbt

from structure import Structure

U = 1000  # divider to get kN


class Importer(Structure):
    """Reads the loads from the model, imports them as a list of pyLoad objects."""

    def __init__(self, app):
        super().__init__(app)

    supported_analize_type = [
        rbt.IRobotCaseAnalizeType.I_CAT_STATIC_LINEAR,
        rbt.IRobotCaseAnalizeType.I_CAT_STATIC_NONLINEAR,
        rbt.IRobotCaseAnalizeType.I_CAT_STATIC_LINEAR_AUXILIARY,
    ]

    # rect = rbt.IRobotLoadRecordType
    supported_load_types = {
        0: "nodal force",  # rect.I_LRT_NODE_FORCE
        5: "uniform load",  # rect.I_LRT_BAR_UNIFORM
        26: "(FE) uniform",  # rect.I_LRT_UNIFORM
        3: "member force",  # rect.I_LRT_BAR_FORCE_CONCENTRATED
        7: "self-weight",  # rect.I_LRT_DEAD
        6: "trapezoidal load (2p)",  # rect.I_LRT_BAR_TRAPEZOIDALE
        69: "(FE) 2 load on edges",  # rect.I_LRT_LINEAR_ON_EDGES
        89: "Body forces",  # rect.I_LRT_BAR_UNIFORM_MASS
    }

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
                    return "MISSING!!!"
            else:
                for r in range(1, count + 1):
                    objects.append(rec.Objects.Get(r))
        else:
            if rec.GetValue(12) == 0:
                objects.append("all")
            else:
                objects.append("objects")
        return objects

    def store_load(self, lcase, rec):
        """Reads the load from the model, stores it in pyLoad object and returns them in list"""
        load = pyLoad()
        rec_type = int(rec.Type)  # cast it as a int
        load.LCName = lcase.Name
        load.objects = self.read_objects(rec)  # read the objects load is assigned to
        if rec_type == 0:  # nodal force
            load.Name = self.supported_load_types[rec_type]
            load.FX = rec.GetValue(0) / U  # I_NFIPRV_FX
            load.FY = rec.GetValue(1) / U  # I_NFIPRV_FY
            load.FZ = rec.GetValue(2) / U  # I_NFIPRV_FZ
            load.Mx = rec.GetValue(3) / U  # I_NFIPRV_MX
            load.My = rec.GetValue(4) / U  # I_NFIPRV_MY
            load.Mz = rec.GetValue(5) / U  # I_NFIPRV_MZ
            load.alfa = rec.GetValue(8)  # I_NFIPRV_ALPHA
            load.beta = rec.GetValue(9)  # I_NFIPRV_BETA
            load.gamma = rec.GetValue(10)  # I_NFIPRV_GAMMA
        elif rec_type == 26:  # uniform load on a FE element
            load.Name = self.supported_load_types[rec_type]
            load.PX = rec.GetValue(0) / U  # I_URV_PX
            load.PY = rec.GetValue(1) / U  # I_URV_PY
            load.PZ = rec.GetValue(2) / U  # I_URV_PZ
            load.cosystem = self.read_cosystem(rec.GetValue(11))  # I_URV_LOCAL_SYSTEM
            load.projected = rec.GetValue(12)  # I_URV_PROJECTION
        elif rec_type == 7:  # self-weight
            load.Name = self.supported_load_types[rec_type]
            load.entirestruc = rec.GetValue(15)  # I_BDRV_ENTIRE_STRUCTURE
        elif rec_type == 3:  # point load on a bar
            load.Name = self.supported_load_types[rec_type]
            load.FX = rec.GetValue(0) / U  # I_BFCRV_FX
            load.FY = rec.GetValue(1) / U  # I_BFCRV_FY
            load.FZ = rec.GetValue(2) / U  # I_BFCRV_FZ
            load.Mx = rec.GetValue(3) / U  # I_BFCRV_CX
            load.My = rec.GetValue(4) / U  # I_BFCRV_CY
            load.Mz = rec.GetValue(5) / U  # I_BFCRV_CZ
            load.alfa = rec.GetValue(8)  # I_BFCRV_ALPHA
            load.beta = rec.GetValue(9)  # I_BFCRV_BETA
            load.gamma = rec.GetValue(10)  # I_BFCRV_GAMMA
            load.cosystem = self.read_cosystem(rec.GetValue(11))  # I_BFCRV_LOC
            load.calcnode = self.read_calcnode(rec.GetValue(12))  # I_BFCRV_GENERATE_CALC_NODE
            load.absrel = self.read_relabs(rec.GetValue(13))  # I_BFCRV_REL
            load.disX = rec.GetValue(6)  # I_BFCRV_X
            load.disY = rec.GetValue(21)  # I_BFCRV_OFFSET_Y
            load.disZ = rec.GetValue(22)  # I_BFCRV_OFFSET_Z
        elif rec_type == 6:
            load.Name = self.supported_load_types[rec_type]
            load.PX = rec.GetValue(3) / U  # I_BTRV_PX1
            load.PY = rec.GetValue(4) / U  # I_BTRV_PY1
            load.PZ = rec.GetValue(5) / U  # I_BTRV_PZ1
            load.PX2 = rec.GetValue(0) / U  # I_BTRV_PX2
            load.PY2 = rec.GetValue(1) / U  # I_BTRV_PY2
            load.PZ2 = rec.GetValue(2) / U  # I_BTRV_PZ2
            load.disX = rec.GetValue(7)  # I_BTRV_X1
            load.disX2 = rec.GetValue(6)  # I_BTRV_X2
            load.alfa = rec.GetValue(8)  # I_BTRV_ALPHA
            load.beta = rec.GetValue(9)  # I_BTRV_BETA
            load.gamma = rec.GetValue(10)  # I_BTRV_GAMMA
            load.projected = rec.GetValue(12)  # I_BTRV_PROJECTION
            load.absrel = self.read_relabs(rec.GetValue(13))  # I_BTRV_RELATIVE
        elif rec_type == 69:
            load.Name = self.supported_load_types[rec_type]
            load.PX = rec.GetValue(0) / U  # I_LOERV_PX
            load.PY = rec.GetValue(1) / U  # I_LOERV_PY
            load.PZ = rec.GetValue(2) / U  # I_LOERV_PZ
            load.Mx = rec.GetValue(3) / U  # I_LOERV_MX
            load.My = rec.GetValue(4) / U  # I_LOERV_MY
            load.Mz = rec.GetValue(5) / U  # I_LOERV_MZ
            load.gamma = rec.GetValue(6)  # I_LOERV_GAMMA
            load.cosystem = self.read_cosystem(rec.GetValue(11))  # I_LOERV_LOCAL_SYSTEM
        elif rec_type == 89:
            load.Name = self.supported_load_types[rec_type]
            load.FX = rec.GetValue(0)
            load.FY = rec.GetValue(1)
            load.FZ = rec.GetValue(2)
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
                    records.append(self.store_load(lcase, rec))
        return records


if __name__ == "__main__":
    app = RobotApplication()
    # pyrobot = pyARSAReporting(app)
    importer = Importer(app)
    records = importer.get_load_records()

    for r in records:
        print(r)

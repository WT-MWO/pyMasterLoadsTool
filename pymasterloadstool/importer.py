import clr
import sys
import math

# from .pyLoad import pyLoad, missing_msg
from openpyxl import Workbook, load_workbook


clr.AddReference(r"C:\Program Files\Autodesk\Robot Structural Analysis Professional 2023\Exe\Interop.RobotOM.dll")
from RobotOM import *
import RobotOM as rbt

from .structure import Structure, supported_load_types, supported_cases_nature, combination_type
from .settings import load_sheet_name, range_to_clear, cases_sheet_name, combinations_sheet_name
from pymasterloadstool import utilities
from .enums import supported_analize_type

U = 1000  # divider to get kN
R = 2  # rounding
rad_to_deg = 180 / math.pi
missing_msg = "MISSING!!"

# TODO: Import loadcase settings nlin etc.


class Importer(Structure):
    """Reads the loads and combinations from the model, imports them as a list of pyLoad objects."""

    def __init__(self, app, path):
        super().__init__(app)
        self.supported_load_types = supported_load_types
        self.path = path
        self.wb = load_workbook(self.path)

    # rect = rbt.IRobotLoadRecordType

    def _read_cosystem(self, cosys):
        if cosys == 0:
            return "global"
        else:
            return "local"

    def _read_calcnode(self, cnode):
        if cnode == 0:
            return "no"
        else:
            return "generated"

    def _read_relabs(self, rela):
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

    def _list_to_str(self, list):
        return ", ".join(repr(e).replace("'", "") for e in list)

    def is_comb_nonlinear(self, lcase):
        if int(lcase.AnalizeType) == -1 or int(lcase.AnalizeType) == -3:
            return 1
        else:
            return 0

    def _read_objects(self, rec):
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

    def _write_load(self, lcase, rec, row):
        """Reads the load from the model, stores it in pyLoad object and returns them in list"""
        # load = pyLoad()
        rec_type = int(rec.Type)  # cast it as a int
        self.ws["B" + str(row)] = lcase.Name
        self.ws["A" + str(row)] = lcase.Number
        self.ws["H" + str(row)] = self.supported_load_types[rec_type]
        self.ws["I" + str(row)] = self._list_to_str(self._read_objects(rec))  # read the objects load is assigned to
        if rec_type == 0 or rec_type == 3:  # nodal force or point load on a bar
            self.ws["J" + str(row)] = round(rec.GetValue(0) / U, R)  # I_NFIPRV_FX
            self.ws["K" + str(row)] = round(rec.GetValue(1) / U, R)  # I_NFIPRV_FY
            self.ws["L" + str(row)] = round(rec.GetValue(2) / U, R)  # I_NFIPRV_FZ
            self.ws["M" + str(row)] = round(rec.GetValue(3) / U, R)  # I_NFIPRV_MX
            self.ws["N" + str(row)] = round(rec.GetValue(4) / U, R)  # I_NFIPRV_MY
            self.ws["O" + str(row)] = round(rec.GetValue(5) / U, R)  # I_NFIPRV_MZ
            self.ws["P" + str(row)] = rec.GetValue(8) * rad_to_deg  # I_NFIPRV_ALPHA
            self.ws["Q" + str(row)] = rec.GetValue(9) * rad_to_deg  # I_NFIPRV_BETA
            self.ws["R" + str(row)] = rec.GetValue(10) * rad_to_deg  # I_NFIPRV_GAMMA
            # load.calcnode = self.read_calcnode(rec.GetValue(12))
        if rec_type == 5:  # uniform load
            self.ws["J" + str(row)] = round(rec.GetValue(0) / U, R)
            self.ws["K" + str(row)] = round(rec.GetValue(1) / U, R)
            self.ws["L" + str(row)] = round(rec.GetValue(2) / U, R)
            self.ws["M" + str(row)] = round(rec.GetValue(3) / U, R)
            self.ws["N" + str(row)] = round(rec.GetValue(4) / U, R)
            self.ws["O" + str(row)] = round(rec.GetValue(5) / U, R)
            self.ws["P" + str(row)] = rec.GetValue(8) * rad_to_deg
            self.ws["Q" + str(row)] = rec.GetValue(9) * rad_to_deg
            self.ws["R" + str(row)] = rec.GetValue(10) * rad_to_deg
            self.ws["W" + str(row)] = self._read_cosystem(rec.GetValue(11))
            # load.projected = rec.GetValue(12)
            self.ws["X" + str(row)] = self._read_relabs(rec.GetValue(13))
            self.ws["T" + str(row)] = rec.GetValue(21)
            self.ws["U" + str(row)] = rec.GetValue(22)
        elif rec_type == 26:  # uniform load on a FE element
            self.ws["J" + str(row)] = round(rec.GetValue(0) / U, R)  # I_URV_PX
            self.ws["K" + str(row)] = round(rec.GetValue(1) / U, R)  # I_URV_PY
            self.ws["L" + str(row)] = round(rec.GetValue(2) / U, R)  # I_URV_PZ
            self.ws["W" + str(row)] = self._read_cosystem(rec.GetValue(11))  # I_URV_LOCAL_SYSTEM
            self.ws["Y" + str(row)] = rec.GetValue(12)  # I_URV_PROJECTION
        elif rec_type == 7:  # self-weight
            self.ws["I" + str(row)] = rec.GetValue(15)  # I_BDRV_ENTIRE_STRUCTURE
        elif rec_type == 3:  # point load on a bar
            self.ws["W" + str(row)] = self._read_cosystem(rec.GetValue(11))  # I_BFCRV_LOC
            self.ws["V" + str(row)] = self._read_calcnode(rec.GetValue(12))  # I_BFCRV_GENERATE_CALC_NODE
            self.ws["X" + str(row)] = self._read_relabs(rec.GetValue(13))  # I_BFCRV_REL
            self.ws["S" + str(row)] = rec.GetValue(6)  # I_BFCRV_X
            self.ws["T" + str(row)] = rec.GetValue(21)  # I_BFCRV_OFFSET_Y
            self.ws["U" + str(row)] = rec.GetValue(22)  # I_BFCRV_OFFSET_Z
        elif rec_type == 6:  # "trapezoidal load (2p)"
            self.ws["M" + str(row)] = round(rec.GetValue(3) / U, R)  # I_BTRV_PX1
            self.ws["N" + str(row)] = round(rec.GetValue(4) / U, R)  # I_BTRV_PY1
            self.ws["O" + str(row)] = round(rec.GetValue(5) / U, R)  # I_BTRV_PZ1
            self.ws["J" + str(row)] = round(rec.GetValue(0) / U, R)  # I_BTRV_PX2
            self.ws["K" + str(row)] = round(rec.GetValue(1) / U, R)  # I_BTRV_PY2
            self.ws["L" + str(row)] = round(rec.GetValue(2) / U, R)  # I_BTRV_PZ2
            self.ws["T" + str(row)] = rec.GetValue(7)  # I_BTRV_X1
            self.ws["S" + str(row)] = rec.GetValue(6)  # I_BTRV_X2
            self.ws["P" + str(row)] = rec.GetValue(8) * rad_to_deg  # I_BTRV_ALPHA
            self.ws["Q" + str(row)] = rec.GetValue(9) * rad_to_deg  # I_BTRV_BETA
            self.ws["R" + str(row)] = rec.GetValue(10) * rad_to_deg  # I_BTRV_GAMMA
            self.ws["Y" + str(row)] = rec.GetValue(12)  # I_BTRV_PROJECTION
            self.ws["X" + str(row)] = self._read_relabs(rec.GetValue(13))  # I_BTRV_RELATIVE
        elif rec_type == 69:  # (FE) 2 load on edges
            self.ws["J" + str(row)] = round(rec.GetValue(0) / U, R)  # I_LOERV_PX
            self.ws["K" + str(row)] = round(rec.GetValue(1) / U, R)  # I_LOERV_PY
            self.ws["L" + str(row)] = round(rec.GetValue(2) / U, R)  # I_LOERV_PZ
            self.ws["M" + str(row)] = round(rec.GetValue(3) / U, R)  # I_LOERV_MX
            self.ws["N" + str(row)] = round(rec.GetValue(4) / U, R)  # I_LOERV_MY
            self.ws["O" + str(row)] = round(rec.GetValue(5) / U, R)  # I_LOERV_MZ
            self.ws["R" + str(row)] = rec.GetValue(6)  # I_LOERV_GAMMA
            self.ws["W" + str(row)] = self._read_cosystem(rec.GetValue(11))  # I_LOERV_LOCAL_SYSTEM
        elif rec_type == 89:  # "Body forces"
            self.ws["J" + str(row)] = round(rec.GetValue(0), R)
            self.ws["K" + str(row)] = round(rec.GetValue(1), R)
            self.ws["L" + str(row)] = round(rec.GetValue(2), R)
            self.ws["X" + str(row)] = self._read_relabs(rec.GetValue(13))  # I_BURV_RELATIVE

    def _import_loadcases(self, cases):
        ws_cases = self.wb[cases_sheet_name]
        row = 7
        for i in range(1, cases.Count + 1):  # loop1
            lcase = rbt.IRobotCase(cases.Get(i))
            if lcase.Type == rbt.IRobotCaseType.I_CT_SIMPLE:
                ws_cases["A" + str(row)] = lcase.Number
                ws_cases["B" + str(row)] = lcase.Name
                ws_cases["C" + str(row)] = supported_cases_nature[int(lcase.Nature)]
                case = rbt.IRobotSimpleCase(lcase)
                if case.IsAuxiliary:
                    ws_cases["I" + str(row)] = 1
                else:
                    ws_cases["I" + str(row)] = 0
                if case.AnalizeType == rbt.IRobotCaseAnalizeType.I_CAT_STATIC_LINEAR:
                    ws_cases["E" + str(row)] = 0
                    ws_cases["F" + str(row)] = 1
                elif lcase.AnalizeType == rbt.IRobotCaseAnalizeType.I_CAT_STATIC_NONLINEAR:
                    ws_cases["E" + str(row)] = 1
                    ws_cases["F" + str(row)] = 2
                    params = rbt.IRobotNonlinearAnalysisParams(case.GetAnalysisParams())
                    if params.MatrixUpdateAfterEachIteration:
                        ws_cases["G" + str(row)] = 1
                    if params.PDelta:
                        ws_cases["H" + str(row)] = 1
                elif lcase.AnalizeType == rbt.IRobotCaseAnalizeType.I_CAT_STATIC_BUCKLING:
                    ws_cases["E" + str(row)] = 0
                    ws_cases["F" + str(row)] = 4
                    params = rbt.IRobotBucklingAnalysisParams(case.GetAnalysisParams())
                    if params.MatrixUpdateAfterEachIteration:
                        ws_cases["G" + str(row)] = 1
                    if params.PDelta:
                        ws_cases["H" + str(row)] = 1
            row += 1

    def _import_combinations(self, cases):
        ws_comb = self.wb[combinations_sheet_name]
        utilities.clear_range(ws_comb, "A7:XFD476")
        row = 7
        for i in range(1, cases.Count + 1):  # loop2
            lcase = rbt.IRobotCase(cases.Get(i))
            # print(int(lcase.Type))
            if int(lcase.Type) == 1:
                # print(lcase.Type)
                lcomb = rbt.IRobotCaseCombination(lcase)
                ws_comb["A" + str(row)] = lcomb.Number
                ws_comb["B" + str(row)] = lcomb.Name
                ws_comb["C" + str(row)] = combination_type[int(lcomb.CombinationType)]
                ws_comb["D" + str(row)] = self.is_comb_nonlinear(lcomb)
                if self.is_comb_nonlinear(lcomb):
                    params = rbt.IRobotNonlinearAnalysisParams(lcomb.GetAnalysisParams())
                    if params.MatrixUpdateAfterEachIteration:
                        ws_comb["F" + str(row)] = 1
                    if params.PDelta:
                        ws_comb["G" + str(row)] = 1
                ws_comb["E" + str(row)] = int(lcomb.AnalizeType)
                row += 1
        # self.wb.save(self.path)

    def _import_load(self, lcase):
        start_row = 8
        if lcase.AnalizeType in supported_analize_type:
            for r in range(1, lcase.Records.Count + 1):
                # Get load record
                rec = lcase.Records.Get(r)
                # Check if load record is not auto generated from claddings loads
                rec_com = rbt.IRobotLoadRecordCommon(lcase.Records.Get(r))
                if not rec_com.IsAutoGenerated:
                    # check if record is supported
                    if self._check_type(rec):
                        self._write_load(lcase=lcase, rec=rec, row=start_row)
                        start_row += 1

    def import_loads_and_comb(self, import_loads=True, import_comb=True):
        "Returns a list of load records of pyLoad object."
        self.ws = self.wb[load_sheet_name]
        # clear the range
        for r in range_to_clear:
            utilities.clear_range(self.ws, r)
        cases = self.structure.Cases.GetAll()
        # Import loadcases
        self._import_loadcases(cases)
        # Import combinations
        if import_comb:
            self._import_combinations(cases)
        for i in range(1, cases.Count + 1):
            lcase = rbt.IRobotCase(cases.Get(i))
            if import_loads:
                self._import_load(lcase)
        self.wb.save(self.path)

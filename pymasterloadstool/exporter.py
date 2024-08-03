import math
import clr

clr.AddReference(r"C:\Program Files\Autodesk\Robot Structural Analysis Professional 2023\Exe\Interop.RobotOM.dll")
from RobotOM import *
import RobotOM as rbt

from .structure import Structure
from .enums import cases_nature, case_analize_type

M = 1000  # multiplier to get N or Pa
deg_to_rad = math.pi / 180


class Exporter(Structure):
    """Exports the loads from the excel to the Robot model"""

    def __init__(self, app):
        super().__init__(app)

    def _del_all_cases(self):
        """Deletes all simple loadcases in the model"""
        case_selection = self.structure.Selections.CreatePredefined(
            rbt.IRobotPredefinedSelection.I_PS_CASE_SIMPLE_CASES
        )
        # all_cases = self.cases.GetAll()
        # for i in range(1, all_cases.Count + 1):
        #     lcase = rbt.IRobotCase(all_cases.Get(i))
        #     selection.Add(lcase)
        self.cases.DeleteMany(case_selection)

    def _del_all_combinations(self):
        """Deletes all combinations in the model"""
        comb_selection = self.structure.Selections.CreatePredefined(
            rbt.IRobotPredefinedSelection.I_PS_CASE_COMBINATIONS
        )
        self.cases.DeleteMany(comb_selection)

    def apply_load_cases(self, cases):
        """Create load cases in the model
        Paremeters:
        cases: list[number, name, nature string]"""
        for c in cases:
            # append 0-number,1-name,2-nature int, 3-nonlin, 4-solver, 5-kmatrix, 6-pdelta
            number = c[0]
            name = c[1]
            nature = cases_nature[c[2]]
            solver = case_analize_type[c[4]]
            case = self.structure.Cases.CreateSimple(number, name, nature, solver)
            # set parameters if solver is non-linear or buckling
            if c[4] == 2:
                params = rbt.IRobotNonlinearAnalysisParams(case.GetAnalysisParams())
                if c[5] is True:
                    params.MatrixUpdateAfterEachIteration = True
                if c[6] is True:
                    params.PDelta = True
                case.SetAnalysisParams(params)
            if c[4] == 4:
                params = rbt.IRobotBucklingAnalysisParams(case.GetAnalysisParams())
                if c[5] is True:
                    params.MatrixUpdateAfterEachIteration = True
                if c[6] is True:
                    params.PDelta = True
                case.SetAnalysisParams(params)

    def _assign_cosystem(self, load):
        if load.cosystem == "global":
            return 0
        else:
            return 1

    def _assign_relabs(self, load):
        if load.absrel == "absolute":
            return 0
        else:
            return 1

    def apply_loads(self, loads):
        for load in loads:
            if load.type == 7:
                case_number = load.LCNumber
                case = rbt.IRobotSimpleCase(self.structure.Cases.Get(case_number))
                record_index = case.Records.New(rbt.IRobotLoadRecordType(7))
                record = rbt.IRobotLoadRecord(case.Records.Get(record_index))
                record.Objects.FromText(load.objects)
                record.SetValue(15, 1)
                record.SetValue(2, -1)
            elif load.type == 0:
                pass
            elif load.type == 5:
                case_number = load.LCNumber
                # create load
                case = rbt.IRobotSimpleCase(self.structure.Cases.Get(case_number))
                record_index = case.Records.New(rbt.IRobotLoadRecordType(0))
                record = case.Records.Get(record_index)
                record.Objects.FromText(load.objects)
                record.SetValue(0, load.FX * M)
                record.SetValue(1, load.FY * M)
                record.SetValue(2, load.FZ * M)
                record.SetValue(3, load.Mx * M)
                record.SetValue(4, load.My * M)
                record.SetValue(5, load.Mz * M)
                record.SetValue(8, load.alfa * deg_to_rad)
                record.SetValue(9, load.beta * deg_to_rad)
                record.SetValue(10, load.gamma * deg_to_rad)
                record.SetValue(11, self._assign_cosystem(load))
                record.SetValue(13, self._assign_relabs(load))
                record.SetValue(21, load.disY)
                record.SetValue(22, load.disZ)

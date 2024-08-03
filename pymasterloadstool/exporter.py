import clr

clr.AddReference(r"C:\Program Files\Autodesk\Robot Structural Analysis Professional 2023\Exe\Interop.RobotOM.dll")
from RobotOM import *
import RobotOM as rbt

from .structure import Structure
from .enums import cases_nature, case_analize_type


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

import clr

clr.AddReference(r"C:\Program Files\Autodesk\Robot Structural Analysis Professional 2023\Exe\Interop.RobotOM.dll")
from RobotOM import *
import RobotOM as rbt

supported_analize_type = [
    rbt.IRobotCaseAnalizeType.I_CAT_STATIC_LINEAR,
    rbt.IRobotCaseAnalizeType.I_CAT_STATIC_NONLINEAR,
    rbt.IRobotCaseAnalizeType.I_CAT_STATIC_LINEAR_AUXILIARY,
]

cases_nature = {
    0: rbt.IRobotCaseNature.I_CN_PERMANENT,
    1: rbt.IRobotCaseNature.I_CN_EXPLOATATION,
    2: rbt.IRobotCaseNature.I_CN_WIND,
    3: rbt.IRobotCaseNature.I_CN_SNOW,
    4: rbt.IRobotCaseNature.I_CN_TEMPERATURE,
    5: rbt.IRobotCaseNature.I_CN_ACCIDENTAL,
    6: rbt.IRobotCaseNature.I_CN_SEISMIC,
}

case_analize_type = {
    0: rbt.IRobotCaseAnalizeType.I_CAT_COMB,
    1: rbt.IRobotCaseAnalizeType.I_CAT_STATIC_LINEAR,
    2: rbt.IRobotCaseAnalizeType.I_CAT_STATIC_NONLINEAR,
    5: rbt.IRobotCaseAnalizeType.I_CAT_STATIC_LINEAR_AUXILIARY,
    11: rbt.IRobotCaseAnalizeType.I_CAT_DYNAMIC_MODAL,
    12: rbt.IRobotCaseAnalizeType.I_CAT_DYNAMIC_SPECTRAL,
    13: rbt.IRobotCaseAnalizeType.I_CAT_DYNAMIC_SEISMIC,
    14: rbt.IRobotCaseAnalizeType.I_CAT_DYNAMIC_HARMONIC,
    30: rbt.IRobotCaseAnalizeType.I_CAT_MOBILE_MAIN,
    4: rbt.IRobotCaseAnalizeType.I_CAT_STATIC_BUCKLING,
    7: rbt.IRobotCaseAnalizeType.I_CAT_STATIC_NONLINEAR_BUCKLING,
}

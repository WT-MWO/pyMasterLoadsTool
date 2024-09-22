class Structure:
    def __init__(self, app):
        """
        This is Python wrapper for Autodesk Robot API.
        Parameters:
            app(RobotApplication): Robot application
        """
        self.app = app
        self.project = self.app.Project
        self.structure = self.project.Structure
        self.labels = self.structure.Labels
        self.cases = self.structure.Cases
        self.nodes = self.structure.Nodes
        self.results = self.structure.Results  # IRobotResultServer
        self.objects = self.structure.Objects


supported_load_types = {
    0: "nodal force",  # rect.I_LRT_NODE_FORCE
    5: "uniform load",  # rect.I_LRT_BAR_UNIFORM
    26: "(FE) uniform",  # rect.I_LRT_UNIFORM
    3: "member force",  # rect.I_LRT_BAR_FORCE_CONCENTRATED
    7: "self-weight",  # rect.I_LRT_DEAD
    6: "trapezoidal load (2p)",  # rect.I_LRT_BAR_TRAPEZOIDALE
    69: "(FE) linear load on edges",  # rect.I_LRT_LINEAR_ON_EDGES
    89: "self-weight",  # body force not supported properly, converted to self-weight
    28: "load on contour",  # I_LRT_IN_CONTOUR
}

supported_cases_nature = {
    0: "dead",
    1: "live",
    2: "wind",
    3: "snow",
    4: "temperature",
    5: "accidental",
    6: "seismic",
}

combination_type = {
    0: "ULS",
    1: "SLS",
    2: "ALS",
    3: "SPC",
}

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
    89: "Body forces",  # rect.I_LRT_BAR_UNIFORM_MASS
}

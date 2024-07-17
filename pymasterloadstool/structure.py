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

import pgdx.report

class ReportGenerator(pgdx.report.ReportGenerator):
    """
    
    """
    def __init__(self, trigger_file, outdir):
        """

        :param trigger_file:
        """
        self._trigger_file = trigger_file
        self._outdir = outdir

    def generateReport(self):
        """

        :return:
        """
        print("Generated report")
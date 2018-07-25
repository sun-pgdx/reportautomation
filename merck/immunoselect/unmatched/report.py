import pgdx.report

import openpyxl

class ReportGenerator(pgdx.report.ReportGenerator):
    """
    
    """
    def __init__(self, trigger_file, outdir):
        """

        :param trigger_file:
        """
        self._trigger_file = trigger_file
        self._outdir = outdir
        self._template_directory = 'templates'
        self._excel_template_name = 'Merck_ImmunoSELECT_Matched_Report_Template.xlsx'
        self._template_file = self._template_directory + '/' + self._excel_template_name
        self._outfile = self._outdir + self._excel_template_name

    def generateReport(self):
        """

        :return:
        """
        xfile = openpyxl.load_workbook(self._template_file)
        self._write_overview_sheet()
        self._write_results_summary_sheet()
        xfile.save(self._outfile)
        print("Wrote output file '%s'" % self._outfile)

    def _write_overview_sheet(self):
        """

        :return:
        """
        sheet = xfile.get_sheet_by_name('Overview')
        sheet['A14'] = "test"

        # pass

    def _write_results_summary_sheet(self):
        """

        :return:
        """
        sheet = xfile.get_sheet_by_name('Results summary')
        sheet['A14'] = "test"

        # pass
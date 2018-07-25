import pgdx.report

import openpyxl

OVERVIEW_CASE_ID = 'A5'
OVERVIEW_DATE = 'C5'
OVERVIEW_TUMOR_TYPE = 'B15'
OVERVIEW_TUMOR_LOCATION = 'B16'
OVERVIEW_SAMPLE_TYPE = 'B17'
OVERVIEW_PATHOLOGICAL_TUMOR_PURITY = 'B18'
OVERVIEW_MUTATION_BASE_TUMOR_PURITY = 'B19'
OVERVIEW_SOURCE_OF_NORMAL_DNA = 'B20'
OVERVIEW_RANDOMIZATION_NUMBER = 'B30'
OVERVIEW_TRIAL_ID = 'B31'
OVERVIEW_SCREENING_NUMBER = 'B32'

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
        self._xfile = openpyxl.load_workbook(self._template_file)

        self._write_overview_sheet()

        self._write_results_summary_sheet()

        self._xfile.save(self._outfile)

        print("Wrote output file '%s'" % self._outfile)

    def _write_overview_sheet(self):
        """

        :return:
        """
        sheet_name = 'Overview'

        sheet = self._xfile.get_sheet_by_name(sheet_name)

        sheet['A14'] = "test"

        print("Wrote to sheet '%s'" % sheet_name)

        # pass

    def _write_results_summary_sheet(self):
        """

        :return:
        """
        sheet_name = 'Results summary'

        sheet = self._xfile.get_sheet_by_name(sheet_name)

        sheet['A14'] = "test"

        print("Wrote to sheet '%s'" % sheet_name)

        # pass
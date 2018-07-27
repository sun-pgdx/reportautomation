import sys

import pgdx.report
from trigger.file.parser import Parser as tp
from summarysheet.file.parser import Parser as ssp
from copy_number.file.parser import Parser as cnp
from final_peptides.file.parser import Parser as fpp
from combined_coverage.file.parser import Parser as ccp
from neoantigens_reported.file.parser import Parser as nrp

from somatic_peptides.xlsx.sheet.writer import Writer as spsw
from copy_number.xlsx.sheet.writer import Writer as cnsw
from somatic_mutations.xlsx.sheet.writer import Writer as smsw
from neoantigen_candidates.xlsx.sheet.writer import Writer as ncsw


import openpyxl

# Overview sheet

OVERVIEW_SHEET_NAME = 'Overview'

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

# Results summary sheet

RESULTS_SUMMARY_SHEET_NAME = 'Results Summary'

RESULTS_SUMMARY_CASE_ID = 'A5'
RESULTS_SUMMARY_DATE = 'C5'

# Number of somatic sequence alterations identified
RESULT_SUMMARY_NUM_SOM_SEQ_ALT_IDE_TUMOR = 'B11'
RESULT_SUMMARY_NUM_SOM_SEQ_ALT_IDE_NORMAL = 'C11'

# Number of somatic copy number alterations identified
RESULT_SUMMARY_NUM_SOM_COP_NUM_ALT_IDE_TUMOR = 'B12'
RESULT_SUMMARY_NUM_SOM_COP_NUM_ALT_IDE_NORMAL = 'C12'

# Sequenced Bases Mapped to Genome
RESULT_SUMMARY_SEQ_BAS_MAP_GEN_TUMOR = 'B15'
RESULT_SUMMARY_SEQ_BAS_MAP_GEN_NORMAL = 'C15'

# Sequenced Bases Mapped to Target Regions
RESULT_SUMMARY_SEQ_BAS_MAP_TAR_REG_TUMOR = 'B16'
RESULT_SUMMARY_SEQ_BAS_MAP_TAR_REG_NORMAL = 'C16'

# Fraction of Sequenced Bases Mapped to Target Regions
RESULT_SUMMARY_FRA_SEQ_BAS_MAP_TAR_REG_TUMOR = 'B17'
RESULT_SUMMARY_FRA_SEQ_BAS_MAP_TAR_REG_NORMAL = 'C17'

# Bases in target regions with at least 10 reads
RESULT_SUMMARY_BAS_TAR_REG_LEA_10_REA_TUMOR = 'B18'
RESULT_SUMMARY_BAS_TAR_REG_LEA_10_REA_NORMAL = 'C18'

# Fraction of bases in target regions with at least 10 reads
RESULT_SUMMARY_FRA_BAS_TAR_REG_LEA_10_REA_TUMOR = 'B19'
RESULT_SUMMARY_FRA_BAS_TAR_REG_LEA_10_REA_NORMAL = 'C19'

# Average Number of Total High Quality Sequences at Each Base
RESULT_SUMMARY_AVE_NUM_TOT_HIG_QUA_SEQ_EAC_BAS_TUMOR = 'B22'
RESULT_SUMMARY_AVE_NUM_TOT_HIG_QUA_SEQ_EAC_BAS_NORMAL = 'C22'

# Average Number of Distinct High Quality Sequences at Each Base
RESULT_SUMMARY_AVE_NUM_DIS_HIG_QUA_SEQ_EAC_BAS_TUMOR = 'B23'
RESULT_SUMMARY_AVE_NUM_DIS_HIG_QUA_SEQ_EAC_BAS_NORMAL = 'C23'

# Germline SNPs present
RESULT_SUMMARY_GER_SNP_PRE_TUMOR = 'B26'
RESULT_SUMMARY_GER_SNP_PRE_NORMAL = 'C26'

# Percent T/N Matching
RESULT_SUMMARY_PRE_TN_MAT_TUMOR = 'B27'
RESULT_SUMMARY_PRE_TN_MAT_NORMAL = 'C27'

class ReportGenerator(pgdx.report.ReportGenerator):
    """

    """
    def __init__(self, trigger_file, outdir):
        """

        :param trigger_file:
        """
        self._trigger_file = trigger_file
        self._outdir = outdir

        self._trigger_file_parser = tp(self._trigger_file)

        # Should be like Merck_ImmunoSELECT_Unmatched_Report_Template.xlsx
        self._template_file = self._trigger_file_parser.getTemplateFilePath()

        self._outfile = self._trigger_file_parser.getFinalReportPath()  + self._trigger_file_parser.getFinalReportName() + '.xlsx'

        self._copy_number_file_parser = cnp(self._trigger_file_parser.getCopyNumberFile())
        self._summarysheet_file_parser = ssp(self._trigger_file_parser.getSummarysheetFile())
        self._combined_coverage_file_parser = ccp(self._trigger_file_parser.getCombinedCoverageFile())


    def generateReport(self):
        """

        :return:
        """

        self._case_id = 'Case ID: ' + self._trigger_file_parser.getPGDXId() + ' - ' + self._trigger_file_parser.getSpecimenNumber()

        self._date = 'Date: ' + str(self._trigger_file_parser.getDate())

        self._xfile = openpyxl.load_workbook(self._template_file)

        self._write_overview_sheet()

        self._write_results_summary_sheet()

        self._write_somatic_mutations_sheet()

        self._write_copy_number_sheet()

        self._write_neoantigen_candidates_sheet()

        self._write_somatic_peptides_sheet()

        self._xfile.save(self._outfile)

        print("Wrote output file '%s'" % self._outfile)

    def _write_overview_sheet(self):
        """

        :return:
        """
        sheet_name = OVERVIEW_SHEET_NAME

        sheet = self._xfile.get_sheet_by_name(sheet_name)

        sheet[OVERVIEW_CASE_ID] = self._case_id
        sheet[OVERVIEW_DATE] = self._date
        sheet[OVERVIEW_TUMOR_TYPE] = self._trigger_file_parser.getDiagnosis()
        sheet[OVERVIEW_TUMOR_LOCATION] = self._trigger_file_parser.getPrimaryTumorSite()
        sheet[OVERVIEW_SAMPLE_TYPE] = self._trigger_file_parser.getSampleType()
        sheet[OVERVIEW_PATHOLOGICAL_TUMOR_PURITY] = self._trigger_file_parser.getPercentTumor()

        # Mutation based Tumor Purity
        # This value is based on the following calculation:
        # (Sum Distinct Mut Reads / Sum Distinct Total Reads)*2*100
        sheet[OVERVIEW_MUTATION_BASE_TUMOR_PURITY] = self._combined_coverage_file_parser.getMutationBaseTumorPurity()

        sheet[OVERVIEW_SOURCE_OF_NORMAL_DNA] = self._trigger_file_parser.getSourceOfNormalDNA()
        sheet[OVERVIEW_RANDOMIZATION_NUMBER] = self._trigger_file_parser.getRandomizationNumber()
        sheet[OVERVIEW_SCREENING_NUMBER] = self._trigger_file_parser.getScreeningNumber()
        sheet[OVERVIEW_TRIAL_ID] = self._trigger_file_parser.getTrialId()

        print("Wrote to sheet '%s'" % sheet_name)

    def _write_results_summary_sheet(self):
        """

        :return:
        """
        sheet_name = RESULTS_SUMMARY_SHEET_NAME

        sheet = self._xfile.get_sheet_by_name(sheet_name)

        sheet[RESULTS_SUMMARY_CASE_ID] = self._case_id

        sheet[RESULTS_SUMMARY_DATE] = self._date

        # Number of somatic sequence alterations identified
        sheet[RESULT_SUMMARY_NUM_SOM_SEQ_ALT_IDE_TUMOR] = self._combined_coverage_file_parser.getRecordCount()
        sheet[RESULT_SUMMARY_NUM_SOM_SEQ_ALT_IDE_NORMAL]= 'N/A'

        # Number of somatic copy number alterations identified
        sheet[RESULT_SUMMARY_NUM_SOM_COP_NUM_ALT_IDE_TUMOR] = self._copy_number_file_parser.getRecordCount()
        sheet[RESULT_SUMMARY_NUM_SOM_COP_NUM_ALT_IDE_NORMAL] = 'N/A'

        # Sequenced Bases Mapped to Genome
        sheet[RESULT_SUMMARY_SEQ_BAS_MAP_GEN_TUMOR] = self._summarysheet_file_parser.getValueByLocation(7,1)

        # Sequenced Bases Mapped to Target Regions
        sheet[RESULT_SUMMARY_SEQ_BAS_MAP_TAR_REG_TUMOR] = self._summarysheet_file_parser.getValueByLocation(9,1)

        # Fraction of Sequenced Bases Mapped to Target Regions
        sheet[RESULT_SUMMARY_FRA_SEQ_BAS_MAP_TAR_REG_TUMOR] = self._summarysheet_file_parser.getValueByLocation(10,1)

        # Bases in target regions with at least 10 reads
        sheet[RESULT_SUMMARY_BAS_TAR_REG_LEA_10_REA_TUMOR] = self._summarysheet_file_parser.getValueByLocation(11,1)

        # Fraction of bases in target regions with at least 10 reads
        sheet[RESULT_SUMMARY_FRA_BAS_TAR_REG_LEA_10_REA_TUMOR] = self._summarysheet_file_parser.getValueByLocation(12,1)

        # Average Number of Total High Quality Sequences at Each Base
        sheet[RESULT_SUMMARY_AVE_NUM_TOT_HIG_QUA_SEQ_EAC_BAS_TUMOR] = self._summarysheet_file_parser.getValueByLocation(18,1)

        # Average Number of Distinct High Quality Sequences at Each Base
        sheet[RESULT_SUMMARY_AVE_NUM_DIS_HIG_QUA_SEQ_EAC_BAS_TUMOR] = self._summarysheet_file_parser.getValueByLocation(23,1)


        print("Wrote to sheet '%s'" % sheet_name)


    def _write_somatic_mutations_sheet(self):
        """

        :return:
        """
        writer = smsw(self._xfile, self._combined_coverage_file_parser, self._case_id, self._date)

        writer.writeSheet()

    def _write_copy_number_sheet(self):
        """

        :return:
        """
        writer = cnsw(self._xfile, self._copy_number_file_parser, self._case_id, self._date)

        writer.writeSheet()

    def _write_neoantigen_candidates_sheet(self):
        """

        :return:
        """
        parser = nrp(self._trigger_file_parser.getNeoantigensReportedFile())

        writer = ncsw(self._xfile, parser, self._case_id, self._date)

        writer.writeSheet()

    def _write_somatic_peptides_sheet(self):
        """

        :return:
        """
        parser = fpp(self._trigger_file_parser.getFinalPeptidesFile())

        writer = spsw(self._xfile, parser, self._case_id, self._date)

        writer.writeSheet()


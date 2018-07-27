import sys

import pgdx.report
from trigger.file.parser import Parser as tp
from summarysheet.file.parser import Parser as ssp
from copy_number.file.parser import Parser as cnp
from final_peptides.file.parser import Parser as fpp
from combined_coverage.file.parser import Parser as ccp
from neoantigens_reported.file.parser import Parser as nrp

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

# Results summary sheet
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

# Copy number sheet
COPY_NUMBER_CASE_ID = 'A5'
COPY_NUMBER_DATE = 'F5'

COPY_NUMBER_GEN_SYM_COL = 'A'
COPY_NUMBER_GEN_DES_COL = 'B'
COPY_NUMBER_GEN_ACC_COL = 'C'
COPY_NUMBER_NUC_POS_COL = 'D'
COPY_NUMBER_FOL_AMP_COL = 'E'
COPY_NUMBER_MUT_TYP_COL = 'F'

COPY_NUMBER_START_ROW = 10

# Somatic Peptides sheet
SOMATIC_PEPTIDES_GEN_SYM = 'A'
SOMATIC_PEPTIDES_MUT_POS = 'B'
SOMATIC_PEPTIDES_MUT_PEP = 'C'

SOMATIC_PEPTIDES_START_ROW = 10

SOMATIC_PEPTIDES_CASE_ID = 'A5'
SOMATIC_PEPTIDES_DATE = 'C5'

# Somatic mutations sheet

SOMATIC_MUTATIONS_CASE_ID = 'A5'
SOMATIC_MUTATIONS_DATE = 'X5'

SOMATIC_MUTATIONS_START_ROW = 10

SOMATIC_MUTATIONS_GEN_SYM = 'A'
SOMATIC_MUTATIONS_GEN_DES = 'B'
SOMATIC_MUTATIONS_TRA_ACC = 'C'
SOMATIC_MUTATIONS_NUC_GEN = 'D'
SOMATIC_MUTATIONS_NUC_COD = 'E'
SOMATIC_MUTATIONS_AMI_ACI = 'F'
SOMATIC_MUTATIONS_EXO = 'G'
SOMATIC_MUTATIONS_MUT_TYP = 'H'
SOMATIC_MUTATIONS_CONSEQ  = 'I'
SOMATIC_MUTATIONS_SEQ_CON = 'J'
SOMATIC_MUTATIONS_MUT_REA = 'K'
SOMATIC_MUTATIONS_CON_INT = 'L'
SOMATIC_MUTATIONS_BIO_REL = 'M'
SOMATIC_MUTATIONS_CLI_REL_GEN = 'N'
SOMATIC_MUTATIONS_PAT_ANA_GO_MOL_FUN = 'O'
SOMATIC_MUTATIONS_PAT_ANA_GO_BIO_FUN = 'P'
SOMATIC_MUTATIONS_PAT_ANA_GO_ADD_INF = 'Q'
SOMATIC_MUTATIONS_REP_SAM_DE_SOM_MUT = 'R'
SOMATIC_MUTATIONS_REP_SAM_SOM_MUT_SAM_AAR = 'S'
SOMATIC_MUTATIONS_REP_SAM_SOM_MUT_NEA_AAR = 'T'
SOMATIC_MUTATIONS_GEN_REP = 'U'
SOMATIC_MUTATIONS_POS_MUT_PRO_DOM = 'V'
SOMATIC_MUTATIONS_POS_MUT_NEA_PRO_DOM = 'W'
SOMATIC_MUTATIONS_CHASM_SCORE = 'X'


# Neoantigen Candidates sheet
NEOANTIGEN_CANDIDATES_START_ROW = 10;

NEOANTIGEN_CANDIDATES_CASE_ID = 'A5'
NEOANTIGEN_CANDIDATES_DATE = 'CD5'

NEOANTIGEN_CANDIDATES_COLUMNS = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ','BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU','BV','BW','BX','BY','BZ','CA','CB','CC','CD','CE','CF','CG','CH','CI','CJ','CK','CL','CM','CN','CO','CP','CQ','CR','CS','CT','CU','CV','CW','CX','CY','CZ']

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

        # Should be like Merck_ImmunoSELECT_Matched_Report_Template.xlsx
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

        print("Will attempt to write the final report file '%s'" %  self._outfile)

        self._xfile.save(self._outfile)

        print("Wrote output file '%s'" % self._outfile)

    def _write_overview_sheet(self):
        """

        :return:
        """
        sheet_name = 'Overview'

        sheet = self._xfile.get_sheet_by_name(sheet_name)

        sheet[OVERVIEW_CASE_ID] = self._case_id
        sheet[OVERVIEW_DATE] = self._date

        # Tumor Type
        sheet[OVERVIEW_TUMOR_TYPE] = self._trigger_file_parser.getDiagnosis()

        # Tumor Location
        sheet[OVERVIEW_TUMOR_LOCATION] = self._trigger_file_parser.getPrimaryTumorSite()

        # Sample Type
        sheet[OVERVIEW_SAMPLE_TYPE] = self._trigger_file_parser.getSampleType()

        # Pathological Tumor Purity
        sheet[OVERVIEW_PATHOLOGICAL_TUMOR_PURITY] = self._trigger_file_parser.getPercentTumor()

        # Mutation based Tumor Purity
        # This value is based on the following calculation:
        # (Sum Distinct Mut Reads / Sum Distinct Total Reads)*2*100
        sheet[OVERVIEW_MUTATION_BASE_TUMOR_PURITY] = self._combined_coverage_file_parser.getMutationBaseTumorPurity()

        # Source of normal DNA
        sheet[OVERVIEW_SOURCE_OF_NORMAL_DNA] = self._trigger_file_parser.getSourceOfNormalDNA()

        # Randomization Number
        sheet[OVERVIEW_RANDOMIZATION_NUMBER] = self._trigger_file_parser.getRandomizationNumber()

        # Screening Number
        sheet[OVERVIEW_SCREENING_NUMBER] = self._trigger_file_parser.getScreeningNumber()

        # Trial ID
        sheet[OVERVIEW_TRIAL_ID] = self._trigger_file_parser.getTrialId()

        print("Wrote to sheet '%s'" % sheet_name)

    def _write_results_summary_sheet(self):
        """

        :return:
        """
        sheet_name = 'Results summary'

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
        sheet[RESULT_SUMMARY_SEQ_BAS_MAP_GEN_NORMAL] = self._summarysheet_file_parser.getValueByLocation(7,2)

        # Sequenced Bases Mapped to Target Regions
        sheet[RESULT_SUMMARY_SEQ_BAS_MAP_TAR_REG_TUMOR] = self._summarysheet_file_parser.getValueByLocation(9,1)
        sheet[RESULT_SUMMARY_SEQ_BAS_MAP_TAR_REG_NORMAL] = self._summarysheet_file_parser.getValueByLocation(9,2)

        # Fraction of Sequenced Bases Mapped to Target Regions
        sheet[RESULT_SUMMARY_FRA_SEQ_BAS_MAP_TAR_REG_TUMOR] = self._summarysheet_file_parser.getValueByLocation(10,1)
        sheet[RESULT_SUMMARY_FRA_SEQ_BAS_MAP_TAR_REG_NORMAL] = self._summarysheet_file_parser.getValueByLocation(10,2)

        # Bases in target regions with at least 10 reads
        sheet[RESULT_SUMMARY_BAS_TAR_REG_LEA_10_REA_TUMOR] = self._summarysheet_file_parser.getValueByLocation(11,1)
        sheet[RESULT_SUMMARY_BAS_TAR_REG_LEA_10_REA_NORMAL] = self._summarysheet_file_parser.getValueByLocation(11,2)

        # Fraction of bases in target regions with at least 10 reads
        sheet[RESULT_SUMMARY_FRA_BAS_TAR_REG_LEA_10_REA_TUMOR] = self._summarysheet_file_parser.getValueByLocation(12,1)
        sheet[RESULT_SUMMARY_FRA_BAS_TAR_REG_LEA_10_REA_NORMAL] = self._summarysheet_file_parser.getValueByLocation(12,2)

        # Average Number of Total High Quality Sequences at Each Base
        sheet[RESULT_SUMMARY_AVE_NUM_TOT_HIG_QUA_SEQ_EAC_BAS_TUMOR] = self._summarysheet_file_parser.getValueByLocation(18,1)
        sheet[RESULT_SUMMARY_AVE_NUM_TOT_HIG_QUA_SEQ_EAC_BAS_NORMAL] = self._summarysheet_file_parser.getValueByLocation(18,2)

        # Average Number of Distinct High Quality Sequences at Each Base
        sheet[RESULT_SUMMARY_AVE_NUM_DIS_HIG_QUA_SEQ_EAC_BAS_TUMOR] = self._summarysheet_file_parser.getValueByLocation(23,1)
        sheet[RESULT_SUMMARY_AVE_NUM_DIS_HIG_QUA_SEQ_EAC_BAS_NORMAL] = self._summarysheet_file_parser.getValueByLocation(23,2)

        # Germline SNPs present
        sheet[RESULT_SUMMARY_GER_SNP_PRE_TUMOR] = self._summarysheet_file_parser.getValueByLocation(33,1)
        sheet[RESULT_SUMMARY_GER_SNP_PRE_NORMAL] = self._summarysheet_file_parser.getValueByLocation(33,2)

        # Percent T/N Matching
        sheet[RESULT_SUMMARY_PRE_TN_MAT_TUMOR] = self._summarysheet_file_parser.getValueByLocation(35,1)
        sheet[RESULT_SUMMARY_PRE_TN_MAT_NORMAL] = 'N/A'

        print("Wrote to sheet '%s'" % sheet_name)


    def _write_somatic_mutations_sheet(self):
        """

        :return:
        """
        sheet_name = 'Somatic mutations'

        sheet = self._xfile.get_sheet_by_name(sheet_name)

        sheet[SOMATIC_MUTATIONS_CASE_ID] = self._case_id

        sheet[SOMATIC_MUTATIONS_DATE] = self._date

        records = self._combined_coverage_file_parser.getSomaticMutationsSheetRecords()

        current_row = SOMATIC_MUTATIONS_START_ROW

        record_ctr = 0

        for record in records:

            a = SOMATIC_MUTATIONS_GEN_SYM + str(current_row)
            b = SOMATIC_MUTATIONS_GEN_DES + str(current_row)
            c = SOMATIC_MUTATIONS_TRA_ACC + str(current_row)
            d = SOMATIC_MUTATIONS_NUC_GEN + str(current_row)
            e = SOMATIC_MUTATIONS_NUC_COD + str(current_row)
            f = SOMATIC_MUTATIONS_AMI_ACI + str(current_row)
            g = SOMATIC_MUTATIONS_EXO + str(current_row)
            h = SOMATIC_MUTATIONS_MUT_TYP + str(current_row)
            i = SOMATIC_MUTATIONS_CONSEQ + str(current_row)
            j = SOMATIC_MUTATIONS_SEQ_CON + str(current_row)
            k = SOMATIC_MUTATIONS_MUT_REA + str(current_row)
            l = SOMATIC_MUTATIONS_CON_INT + str(current_row)
            m = SOMATIC_MUTATIONS_BIO_REL + str(current_row)
            n = SOMATIC_MUTATIONS_CLI_REL_GEN + str(current_row)
            o = SOMATIC_MUTATIONS_PAT_ANA_GO_MOL_FUN + str(current_row)
            p = SOMATIC_MUTATIONS_PAT_ANA_GO_BIO_FUN + str(current_row)
            q = SOMATIC_MUTATIONS_PAT_ANA_GO_ADD_INF + str(current_row)
            r = SOMATIC_MUTATIONS_REP_SAM_DE_SOM_MUT + str(current_row)
            s = SOMATIC_MUTATIONS_REP_SAM_SOM_MUT_SAM_AAR + str(current_row)
            t = SOMATIC_MUTATIONS_REP_SAM_SOM_MUT_NEA_AAR + str(current_row)
            u = SOMATIC_MUTATIONS_GEN_REP + str(current_row)
            v = SOMATIC_MUTATIONS_POS_MUT_PRO_DOM + str(current_row)
            w = SOMATIC_MUTATIONS_POS_MUT_NEA_PRO_DOM + str(current_row)
            x = SOMATIC_MUTATIONS_CHASM_SCORE + str(current_row)

            sheet[a] = record[0]
            sheet[b] = record[1]
            sheet[c] = record[2]
            sheet[d] = record[3]
            sheet[e] = record[4]
            sheet[f] = record[5]
            sheet[g] = record[6]
            sheet[h] = record[7]
            sheet[i] = record[8]
            sheet[j] = record[9]
            sheet[k] = record[10]
            sheet[l] = record[11]
            sheet[m] = record[12]
            sheet[n] = record[13]
            sheet[o] = record[14]
            sheet[p] = record[15]
            sheet[q] = record[16]
            sheet[r] = record[17]
            sheet[s] = record[18]
            sheet[t] = record[19]
            sheet[u] = record[20]
            sheet[v] = record[21]
            sheet[w] = record[22]
            sheet[x] = record[23]

            current_row += 1

            record_ctr += 1

        print("Wrote '%d' records to sheet '%s'" % (record_ctr, sheet_name))


    def _write_copy_number_sheet(self):
        """

        :return:
        """
        sheet_name = 'Copy number'

        sheet = self._xfile.get_sheet_by_name(sheet_name)

        sheet[COPY_NUMBER_CASE_ID] = self._case_id

        sheet[COPY_NUMBER_DATE] = self._date

        records = self._copy_number_file_parser.getCopyNumberSheetRecords()

        current_row =  COPY_NUMBER_START_ROW

        record_ctr = 0

        for record in records:

            if record_ctr != 0:
                # do not want to write the header of the data file to this sheet

                a = COPY_NUMBER_GEN_SYM_COL + str(current_row)
                b = COPY_NUMBER_GEN_DES_COL + str(current_row)
                c = COPY_NUMBER_GEN_ACC_COL + str(current_row)
                d = COPY_NUMBER_NUC_POS_COL + str(current_row)
                e = COPY_NUMBER_FOL_AMP_COL + str(current_row)
                f = COPY_NUMBER_MUT_TYP_COL + str(current_row)

                sheet[a] = record[0]
                sheet[b] = record[1]
                sheet[c] = record[2]
                sheet[d] = record[3]
                sheet[e] = record[4]
                sheet[f] = record[5]

                current_row += 1

            record_ctr += 1

        if record_ctr == 1:
            sheet['A10'] = 'No copy number alterations identified'


        print("Wrote '%d' records to sheet '%s'" % (record_ctr, sheet_name))



    def _write_neoantigen_candidates_sheet(self):
        """

        :return:
        """

        sheet_name = 'Neoantigen Candidates'

        sheet = self._xfile.get_sheet_by_name(sheet_name)

        sheet[NEOANTIGEN_CANDIDATES_CASE_ID] = self._case_id

        sheet[NEOANTIGEN_CANDIDATES_DATE] = self._date

        self._neoantigens_reported_file_parser = nrp(self._trigger_file_parser.getNeoantigensReportedFile())

        record_list = self._neoantigens_reported_file_parser.getRecordList()

        current_row = NEOANTIGEN_CANDIDATES_START_ROW

        record_ctr = 0

        for record in record_list:

            if record_ctr != 0:
                # do not want to write the header of the data file to this sheet

                record.pop(0) # remove the first element which is the PGDX identifier

                # num_fields = len(record)
                field_ctr = 0
                for field in record:
                    location = NEOANTIGEN_CANDIDATES_COLUMNS[field_ctr]
                    cell_location = location + str(current_row)
                    sheet[cell_location] = field
                    field_ctr += 1

                current_row += 1

            record_ctr += 1

        print("Wrote '%d' records to sheet '%s'" % (record_ctr, sheet_name))



    def _write_somatic_peptides_sheet(self):
        """

        :return:
        """


        sheet_name = 'Somatic Peptides'

        sheet = self._xfile.get_sheet_by_name(sheet_name)

        sheet[SOMATIC_PEPTIDES_CASE_ID] = self._case_id

        sheet[SOMATIC_PEPTIDES_DATE] = self._date

        self._final_peptides_file_parser = fpp(self._trigger_file_parser.getFinalPeptidesFile())

        records = self._final_peptides_file_parser.getSomaticPeptidesSheetRecords()

        current_row = SOMATIC_PEPTIDES_START_ROW

        record_ctr = 0

        for record in records:

            if record_ctr != 0:
                # do not want to write the header of the data file to this sheet

                a = SOMATIC_PEPTIDES_GEN_SYM + str(current_row)
                b = SOMATIC_PEPTIDES_MUT_POS + str(current_row)
                c = SOMATIC_PEPTIDES_MUT_PEP + str(current_row)

                sheet[a] = record[0]
                sheet[b] = record[1]
                sheet[c] = record[2]

                current_row += 1

            record_ctr += 1

        print("Wrote '%d' records to sheet '%s'" % (record_ctr, sheet_name))



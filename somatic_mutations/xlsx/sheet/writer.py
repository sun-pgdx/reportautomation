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


class Writer(object):
    """

    """
    def __init__(self, xfile, file_parser, case_id, date):
        """

        :param xfile:
        :param file_parser:
        :param case_id:
        :param date:
        """
        self._xfile = xfile
        self._file_parser = file_parser
        self._case_id = case_id
        self._date = date
        self._sheet_name = 'Somatic Mutations'

    def writeSheet(self):
        """

        :return:
        """

        sheet = self._xfile.get_sheet_by_name(self._sheet_name)

        sheet[SOMATIC_MUTATIONS_CASE_ID] = self._case_id

        sheet[SOMATIC_MUTATIONS_DATE] = self._date

        records = self._file_parser.getSomaticMutationsSheetRecords()

        current_row = SOMATIC_MUTATIONS_START_ROW

        record_ctr = 0

        for record in records:

            sheet[SOMATIC_MUTATIONS_GEN_SYM + str(current_row)] = record[0]
            sheet[SOMATIC_MUTATIONS_GEN_DES + str(current_row)] = record[1]
            sheet[SOMATIC_MUTATIONS_TRA_ACC + str(current_row)] = record[2]
            sheet[SOMATIC_MUTATIONS_NUC_GEN + str(current_row)] = record[3]
            sheet[SOMATIC_MUTATIONS_NUC_COD + str(current_row)] = record[4]
            sheet[SOMATIC_MUTATIONS_AMI_ACI + str(current_row)] = record[5]
            sheet[SOMATIC_MUTATIONS_EXO + str(current_row)] = record[6]
            sheet[SOMATIC_MUTATIONS_MUT_TYP + str(current_row)] = record[7]
            sheet[SOMATIC_MUTATIONS_CONSEQ + str(current_row)] = record[8]
            sheet[SOMATIC_MUTATIONS_SEQ_CON + str(current_row)] = record[9]
            sheet[SOMATIC_MUTATIONS_MUT_REA + str(current_row)] = record[10]
            sheet[SOMATIC_MUTATIONS_CON_INT + str(current_row)] = record[11]
            sheet[SOMATIC_MUTATIONS_BIO_REL + str(current_row)] = record[12]
            sheet[SOMATIC_MUTATIONS_CLI_REL_GEN + str(current_row)] = record[13]
            sheet[SOMATIC_MUTATIONS_PAT_ANA_GO_MOL_FUN + str(current_row)] = record[14]
            sheet[SOMATIC_MUTATIONS_PAT_ANA_GO_BIO_FUN + str(current_row)] = record[15]
            sheet[SOMATIC_MUTATIONS_PAT_ANA_GO_ADD_INF + str(current_row)] = record[16]
            sheet[SOMATIC_MUTATIONS_REP_SAM_DE_SOM_MUT + str(current_row)] = record[17]
            sheet[SOMATIC_MUTATIONS_REP_SAM_SOM_MUT_SAM_AAR + str(current_row)] = record[18]
            sheet[SOMATIC_MUTATIONS_REP_SAM_SOM_MUT_NEA_AAR + str(current_row)] = record[19]
            sheet[SOMATIC_MUTATIONS_GEN_REP + str(current_row)] = record[20]
            sheet[SOMATIC_MUTATIONS_POS_MUT_PRO_DOM + str(current_row)] = record[21]
            sheet[SOMATIC_MUTATIONS_POS_MUT_NEA_PRO_DOM + str(current_row)] = record[22]
            sheet[SOMATIC_MUTATIONS_CHASM_SCORE + str(current_row)] = record[23]

            current_row += 1

            record_ctr += 1

        print("Wrote '%d' records to sheet '%s'" % (record_ctr, self._sheet_name))

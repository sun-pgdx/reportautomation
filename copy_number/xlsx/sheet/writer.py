COPY_NUMBER_CASE_ID = 'A5'
COPY_NUMBER_DATE = 'F5'

COPY_NUMBER_GEN_SYM_COL = 'A'
COPY_NUMBER_GEN_DES_COL = 'B'
COPY_NUMBER_GEN_ACC_COL = 'C'
COPY_NUMBER_NUC_POS_COL = 'D'
COPY_NUMBER_FOL_AMP_COL = 'E'
COPY_NUMBER_MUT_TYP_COL = 'F'

COPY_NUMBER_START_ROW = 10

class Writer(object):
    """

    """
    def __init__(self, xfile, copy_number_file_parser, case_id, date):
        """

        :param xfile:
        :param copy_number_file_parser:
        :param case_id:
        :param date:
        """
        self._xfile = xfile
        self._copy_number_file_parser = copy_number_file_parser
        self._case_id = case_id
        self._date = date
        self._sheet_name = 'Copy Number'

    def writeSheet(self):
        """

        :return:
        """

        sheet = self._xfile.get_sheet_by_name(self._sheet_name)

        sheet[COPY_NUMBER_CASE_ID] = self._case_id

        sheet[COPY_NUMBER_DATE] = self._date

        records = self._copy_number_file_parser.getCopyNumberSheetRecords()

        current_row = COPY_NUMBER_START_ROW

        record_ctr = 0

        for record in records:

            if record_ctr != 0:
                # do not want to write the header of the data file to this sheet

                sheet[COPY_NUMBER_GEN_SYM_COL + str(current_row)] = record[0]
                sheet[COPY_NUMBER_GEN_DES_COL + str(current_row)] = record[1]
                sheet[COPY_NUMBER_GEN_ACC_COL + str(current_row)] = record[2]
                sheet[COPY_NUMBER_NUC_POS_COL + str(current_row)] = record[3]
                sheet[COPY_NUMBER_FOL_AMP_COL + str(current_row)] = record[4]
                sheet[COPY_NUMBER_MUT_TYP_COL + str(current_row)] = record[5]

                current_row += 1

            record_ctr += 1

        if record_ctr == 1:
            sheet['A10'] = 'No copy number alterations identified'


        print("Wrote '%d' records to sheet '%s'" % (record_ctr, self._sheet_name))

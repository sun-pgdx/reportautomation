# Somatic Peptides sheet
SOMATIC_PEPTIDES_GEN_SYM = 'A'
SOMATIC_PEPTIDES_MUT_POS = 'B'
SOMATIC_PEPTIDES_MUT_PEP = 'C'

SOMATIC_PEPTIDES_START_ROW = 10

SOMATIC_PEPTIDES_CASE_ID = 'A5'
SOMATIC_PEPTIDES_DATE = 'C5'


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
        self._sheet_name = 'Somatic Peptides'


    def writeSheet(self):
        """

        :return:
        """


        sheet = self._xfile.get_sheet_by_name(self._sheet_name)

        sheet[SOMATIC_PEPTIDES_CASE_ID] = self._case_id

        sheet[SOMATIC_PEPTIDES_DATE] = self._date


        records = self._file_parser.getSomaticPeptidesSheetRecords()

        current_row = SOMATIC_PEPTIDES_START_ROW

        record_ctr = 0

        for record in records:

            if record_ctr != 0:
                # do not want to write the header of the data file to this sheet

                sheet[SOMATIC_PEPTIDES_GEN_SYM + str(current_row)] = record[0]
                sheet[SOMATIC_PEPTIDES_MUT_POS + str(current_row)] = record[1]
                sheet[SOMATIC_PEPTIDES_MUT_PEP + str(current_row)] = record[2]

                current_row += 1

            record_ctr += 1

        print("Wrote '%d' records to sheet '%s'" % (record_ctr, self._sheet_name))



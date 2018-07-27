NEOANTIGEN_CANDIDATES_START_ROW = 10;

NEOANTIGEN_CANDIDATES_CASE_ID = 'A5'
NEOANTIGEN_CANDIDATES_DATE = 'CD5'

NEOANTIGEN_CANDIDATES_COLUMNS = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ','BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU','BV','BW','BX','BY','BZ','CA','CB','CC','CD','CE','CF','CG','CH','CI','CJ','CK','CL','CM','CN','CO','CP','CQ','CR','CS','CT','CU','CV','CW','CX','CY','CZ']


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
        self._sheet_name = 'Neoantigen Candidates'

    def writeSheet(self):
        """

        :return:
        """

        sheet = self._xfile.get_sheet_by_name(self._sheet_name)

        sheet[NEOANTIGEN_CANDIDATES_CASE_ID] = self._case_id

        sheet[NEOANTIGEN_CANDIDATES_DATE] = self._date

        record_list = self._file_parser.getRecordList()

        current_row = NEOANTIGEN_CANDIDATES_START_ROW

        record_ctr = 0

        for record in record_list:

            if record_ctr != 0:
                # do not want to write the header of the data file to this sheet

                record.pop(0)  # remove the first element which is the PGDX identifier

                # num_fields = len(record)
                field_ctr = 0

                for field in record:

                    location = NEOANTIGEN_CANDIDATES_COLUMNS[field_ctr]

                    cell_location = location + str(current_row)

                    sheet[cell_location] = field

                    field_ctr += 1

                current_row += 1

            record_ctr += 1

        print("Wrote '%d' records to sheet '%s'" % (record_ctr, self._sheet_name))


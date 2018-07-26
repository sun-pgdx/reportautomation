import pgdx.file.parser

class Parser(pgdx.file.parser.Parser):
    """

    """
    def __init__(self, infile):
        """

        :param infile:
        """
        self._infile = infile
        self._record_list = []
        self._record_count = 0
        self._has_header_row = True
        self._parse_file()


    def getValueByLocation(self, row_num, column_num):
        """

        :param row_num:
        :param column_num:
        :return:
        """
        row = self._record_list[row_num]

        val = row[column_num]

        return val


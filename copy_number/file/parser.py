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


    def getCopyNumberSheetRecords(self):
        """

        :return: list of list
        """
        return self._record_list
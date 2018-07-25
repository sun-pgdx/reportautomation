import csv
import pprint
pp = pprint.PrettyPrinter(indent=4)

class Parser(object):
    """
    """
    def __init__(self, infile):
        """

        :param infile:
        """
        self._infile = infile
        self._lookup = {}
        self._position_to_header_lookup = {}
        self._parse_file()


    def _parse_file(self):
        """

        :return:
        """
        with open(self._infile, 'r') as csvfile:
            reader = csv.reader(csvfile)
            row_ctr= 0
            for row in reader:
                row_ctr += 1
                if row_ctr == 1:
                    field_ctr = 0
                    for field in row:
                        self._position_to_header_lookup[field_ctr] = field
                        field_ctr += 1
                    pp.pprint(self._position_to_header_lookup)
                    print("Processed the header row for trigger file '%s'" % self._infile)
                else:
                    field_ctr = 0
                    for field in row:
                        header = self._position_to_header_lookup[field_ctr]
                        self._lookup[header] = field
                        field_ctr += 1
            print("Processed '%d' rows in trigger file" % row_ctr)

    def getReportType(self):
        """

        :return: self._report_type (string)
        """
        pp.pprint(self._lookup)
        return self._lookup['report_type']
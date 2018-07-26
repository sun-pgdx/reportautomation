import csv

class Parser():
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

    def _parse_file(self):
        """

        :return:
        """

        print("Will attempt to parser file '%s'" % self._infile)

        with open(self._infile, 'r') as csvfile:
            reader = csv.reader(csvfile, delimiter='\t')
            for row in reader:
                self._record_list.append(row)
                self._record_count += 1




    def getRecordCount(self):
        """

        :return:
        """
        if self._has_header_row:
            count = self._record_count - 1
            return count
        else:
            return self._record_count

    def getRecordList(self):
        """

        :return:
        """
        return self._record_list
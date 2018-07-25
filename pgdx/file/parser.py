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
        self._parse_file()

    def _parse_file(self):
        """

        :return:
        """
        with open(self._infile, 'r') as csvfile:
            reader = csv.reader(csvfile, delimiter='\t')
            for row in reader:
                self._record_list.append(row)
                self._record_count += 1


        print("Finished parsing file '%s'" % self._infile)

    def getRecordCount(self):
        """

        :return:
        """
        return self._record_count

    def getRecordList(self):
        """

        :return:
        """
        return self._record_list
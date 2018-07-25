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
                    # pp.pprint(self._position_to_header_lookup)
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
        # pp.pprint(self._lookup)
        return self._lookup['report_type']

    def getClientName(self):
        """

        :return:
        """
        return self._lookup['Client Name']

    def getDiagnosis(self):
        """

        :return:
        """
        return self._lookup['Diagnosis']

    def getPatientMedicalRecord(self):
        """

        :return:
        """
        return self._lookup['Patient Medical Record']

    def getPatientName(self):
        """

        :return:
        """
        return self._lookup['Patient Name']

    def getPercentTumor(self):
        """

        :return:
        """
        return self._lookup['Percent Tumor']

    def getPrimaryTumorSite(self):
        """

        :return:
        """
        return self._lookup['Primar Tumor Site']

    def getProjectName(self):
        """

        :return:
        """
        return self._lookup['Project Name']

    def getProtocol(self):
        """

        :return:
        """
        return self._lookup['Protocol']

    def getSampleType(self):
        """

        :return:
        """
        return self._lookup['Sample Type']

    def getSpecimenNumber(self):
        """

        :return:
        """
        return self._lookup['Specimen #']

    def getTestDisposition(self):
        """

        :return:
        """
        return self._lookup['Test Disposition']


    def getTestsOrdered(self):
        """

        :return:
        """
        return self._lookup['Tests Ordered']

    def getPGDXId(self):
        """

        :return:
        """
        return self._lookup['PGDXID']

    def getSummarysheetFile(self):
        """

        :return:
        """
        return self._lookup['summarysheet']

    def getNeoantigensReportedFile(self):
        """

        :return:
        """
        return self._lookup['neoantigens_reported']

    def getFinalPeptidesFile(self):
        """

        :return:
        """
        return self._lookup['final_peptides']

    def getCombinedCoverageFile(self):
        """

        :return:
        """
        return self._lookup['combined_coverage']

    def getCopyNumberFile(self):
        """

        :return:
        """
        return self._lookup['copy_number']

    def getFinalReportPath(self):
        """

        :return:
        """
        return self._lookup['Final Report Path']

    def getFinalReportName(self):
        """

        :return:
        """
        return self._lookup['Final Report Name']

    def getTemplateFilePath(self):
        """

        :return:
        """
        return self._lookup['Template File Path']

    def getReportId(self):
        """

        :return:
        """
        return self._lookup['Report ID']

    def getRandomizationNumber(self):
        """

        :return:
        """
        return self._lookup['Randomization ID']

    def getScreeningNumber(self):
        """

        :return:
        """
        return self._lookup['Screening ID']

    def getTrialId(self):
        """

        :return:
        """
        return self._lookup['Trial ID']

    def getSourceOfNormalDNA(self):
        """

        :return:
        """
        if 'Source of Normal DNA' in self._lookup:
            return self._lookup['Source of Normal DNA']
        else:
            return 'N/A'



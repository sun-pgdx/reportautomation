import csv
import pgdx.file.parser
import pprint
pp = pprint.PrettyPrinter(indent=4)

# qualified_positions = set((227, 228, 229, 230,231: 'Report: CodonChange',
# 232: 'Report: AAChange',
# 233: 'Report: Exon Rank',
# 234: 'Report: ReportedMutation',
# 235: 'Report: ReportedConsequence',
# 236: 'Report: Seq window',
# 237: 'Report: MutPct',
# 238: 'Report: 95% Confidence Interval for % Mutant Reads',
# 244: 'Report: Biologically Significant',
# 245: 'Report: Clinically Significant',
# 246: 'Report: Pathway Analysis (GO Molecular Function)',
# 247: 'Report: Pathway Analysis (GO Biological Process)',
# 248: 'Report: Pathway Analysis (Additional Information)',
# 249: 'Report: Samples with the Identical Somatic Mutation',
# 251: 'Report: Samples with Somatic Mutations in Nearby Amino Acid Residues',
# 252: 'Report: Gene Reported to be Somatically Mutated in the Following '
# 'Cancers',
# 253: 'Report: Position of Mutation Within a Protein Domain',
# 254: 'Report: Position of Mutation Near to a Protein Domain',
# 255: 'Report: CHASM'
# }


# The following two fields are required for determining the
# Mutation based Tumor Purity where
# 'Report: DistinctPairs' is column number 242 (IH) maps to Sum Distinct Mut Reads
# 'Report: DistinctCoverage' is column number 243 (II) maps to Sum Distinct Total Reads
#
# and is calculated with this formula:
# (Sum Distinct Mut Reads / Sum Distinct Total Reads)*2*100

SUM_DISTINCT_MUT_READS = 241
SUM_DISTINCT_TOTAL_READS = 242


qualified_fields = set(('Report: GeneName',
                        'Report: Description',
                        'Report: Transcript',
                        'Report: LookupKey',
                        'Report: CodonChange',
                        'Report: AAChange',
                        'Report: Exon Rank',
                        'Report: ReportedMutation',
                        'Report: ReportedConsequence',
                        'Report: Seq window',
                        'Report: MutPct',
                        'Report: 95% Confidence Interval for % Mutant Reads',
                        'Report: Biologically Significant',
                        'Report: Clinically Significant',
                        'Report: Pathway Analysis (GO Molecular Function)',
                        'Report: Pathway Analysis (GO Biological Process)',
                        'Report: Pathway Analysis (Additional Information)',
                        'Report: Samples with the Identical Somatic Mutation',
                        'Report: Samples with Somatic Mutations in the Same Amino Acid Residue',
                        'Report: Samples with Somatic Mutations in Nearby Amino Acid Residues',
                        'Report: Gene Reported to be Somatically Mutated in the Following Cancers',
                        'Report: Position of Mutation Within a Protein Domain',
                        'Report: Position of Mutation Near to a Protein Domain',
                        'Report: CHASM'))

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
        self._position_to_header_lookup = {}
        self._somatic_mutations_record_list = []
        self._sum_distinct_mut_reads = 0
        self._sum_distinct_total_reads = 0

        self._parse_file()

    def _parse_file(self):
        """

        :return:
        """

        with open(self._infile, 'r') as csvfile:
            reader = csv.reader(csvfile, delimiter="\t")

            row_ctr= 0

            for row in reader:

                row_ctr += 1

                if row_ctr == 1:
                    # Processing the header
                    header_ctr = 0
                    for header in row:
                        if header in qualified_fields:
                            self._position_to_header_lookup[header_ctr] = header
                        header_ctr += 1
                    print("Processed the header row for trigger file '%s'" % self._infile)
                    #pp.pprint(self._position_to_header_lookup)
                else:
                    # Processing a non-header row

                    smr= []
                    smr.append(row[227].replace("\n", '').replace("\r",''))
                    smr.append(row[228].replace("\n", '').replace("\r",''))
                    smr.append(row[229].replace("\n", '').replace("\r",''))
                    smr.append(row[230].replace("\n", '').replace("\r",''))
                    smr.append(row[231].replace("\n", '').replace("\r",''))
                    smr.append(row[232].replace("\n", '').replace("\r",''))
                    smr.append(row[233].replace("\n", '').replace("\r",''))
                    smr.append(row[234].replace("\n", '').replace("\r",''))
                    smr.append(row[235].replace("\n", '').replace("\r",''))
                    smr.append(row[236].replace("\n", '').replace("\r",''))
                    smr.append(row[237].replace("\n", '').replace("\r",''))
                    smr.append(row[238].replace("\n", '').replace("\r",''))
                    smr.append(row[244].replace("\n", '').replace("\r",''))
                    smr.append(row[245].replace("\n", '').replace("\r",''))
                    smr.append(row[246].replace("\n", '').replace("\r",''))
                    #smr.append('N/A')
                    smr.append(row[247].replace("\n", '').replace("\r",''))
                    #smr.append('N/A')
                    smr.append(row[248].replace("\n", '').replace("\r",''))
                    #smr.append('N/A')
                    smr.append(row[249].replace("\n", '').replace("\r",''))
                    smr.append(row[250].replace("\n", '').replace("\r",''))
                    smr.append(row[251].replace("\n", '').replace("\r",''))
                    smr.append(row[252].replace("\n", '').replace("\r",''))
                    # smr.append('N/A')
                    smr.append(row[253].replace("\n", '').replace("\r",''))
                    smr.append(row[254].replace("\n", '').replace("\r",''))
                    smr.append(row[255].replace("\n", '').replace("\r",''))



                    # IH: report distinctcoverage => Sum Distinct Total Reads
                    # II: Report: DistinctPairs => Sum Distinct Mut Reads
                    # (Sum Distinct Mut Reads / Sum Distinct Total Reads)*2*100

                    self._sum_distinct_mut_reads += float(row[SUM_DISTINCT_MUT_READS])

                    self._sum_distinct_total_reads += float(row[SUM_DISTINCT_TOTAL_READS])

                    # field_ctr = 0
                    # for field in row:
                    #     if field_ctr in self._position_to_header_lookup:
                    #         # this is a field we're interested in
                    #         somatic_mutation_record.append(field)
                    #
                    #     field_ctr += 1

                    self._somatic_mutations_record_list.append(smr)

            # pp.pprint(self._somatic_mutations_record_list)
            print("Processed '%d' rows in combined coverage file '%s'" % (row_ctr, self._infile))



    def getSomaticMutationsSheetRecords(self):
        """

        :return: list of list
        """
        return self._somatic_mutations_record_list

    def getMutationBaseTumorPurity(self):
        """

        :return: list of list
        """
        val = (self._sum_distinct_total_reads / self._sum_distinct_mut_reads) * 200

        return val
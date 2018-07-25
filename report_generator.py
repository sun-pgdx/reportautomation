import os
import sys
import click
import pathlib

from trigger.file.parser import Parser as tp
from copy_number.file.parser import Parser as cnp
from combined_coverage.file.parser import Parser as ccp
from final_peptides.file.parser import Parser as fpp
from neoantigens_reported.file.parser import Parser as nrp
from summarysheet.file.parser import Parser as ssp

from merck.immunoselect.matched.report import ReportGenerator as mimRepGen
from merck.immunoselect.unmatched.report import ReportGenerator as miumRepGen
from victor.immunoselect.matched.report import ReportGenerator as vimRepGen
from victor.immunoselect.unmatched.report import ReportGenerator as viumRepGen

@click.command()
@click.argument('trigger_file')
@click.option('--verbose', default=False, is_flag=True, help='Will print more info to STDOUT')
@click.option('--outdir', default='./', help='The default is the current working directory')
def main(trigger_file, verbose, outdir):
    """

    TRIGGER_FILE: The input file to be processed.

    :return: None


    """
    assert isinstance(trigger_file, str)

    if not os.path.isfile(trigger_file):
        print("'%s' is not a file" % trigger_file)
        sys.exit(1)

    if verbose:
        print("The input file is %s" % trigger_file)

    if not outdir == './':
        pathlib.Path(outdir).mkdir(parents=True, exist_ok=True)
        if verbose:
            print("output directory '%s' was created" % outdir)

    trigger_file_parser = tp(trigger_file)

    report_type = trigger_file_parser.getReportType()

    repgen = None

    if report_type == 'Merck_ImmunoSELECT_Matched' or report_type == 'Merck ImmunoSELECT Matched':
        repgen = mimRepGen(trigger_file, outdir)

    elif report_type == 'Merck_ImmunoSELECT_Unmatched' or report_type == 'Merck ImmunoSELECT Unmatched':
        repgen = miumRepGen(trigger_file, outdir)

    elif report_type == 'Victor ImmunoSELECT Matched':
        print("A1")
        repgen = vimRepGen(trigger_file, outdir)

    elif report_type == 'Victor ImmunoSELECT Unmatched':
        print("A2")
        repgen = viumRepGen(trigger_file, outdir)

    else:
        print("report type '%s' is not supported" % report_type)
        sys.exit(1)

    repgen.generateReport()

if __name__ == "__main__":
    main()

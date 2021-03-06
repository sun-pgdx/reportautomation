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

from immunoselect.matched.report import ReportGenerator as imRepGen
from immunoselect.unmatched.report import ReportGenerator as iumRepGen

from cancerxome.matched.report import ReportGenerator as cxmRepGen
from cancerxome.unmatched.report import ReportGenerator as cxumRepGen

from victor.cancerxome.matched.report import ReportGenerator as vcxmRepGen
from victor.cancerxome.unmatched.report import ReportGenerator as vcxumRepGen


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

    elif report_type == 'Victor ImmunoSELECT Matched' or report_type == 'Victor_ImmunoSELECT_Matched':

        repgen = vimRepGen(trigger_file, outdir)

    elif report_type == 'Victor ImmunoSELECT Unmatched' or report_type == 'Victor_ImmunoSELECT_Unmatched':

        repgen = viumRepGen(trigger_file, outdir)

    elif report_type == 'ImmunoSELECT Matched' or report_type == 'ImmunoSELECT_Matched':

        repgen = imRepGen(trigger_file, outdir)

    elif report_type == 'ImmunoSELECT Unmatched' or report_type == 'ImmunoSELECT_Unmatched':

        repgen = iumRepGen(trigger_file, outdir)

    elif report_type == 'CancerXOME Matched' or report_type == 'CancerXOME_Matched':

        repgen = cxmRepGen(trigger_file, outdir)

    elif report_type == 'CancerXOME Unmatched' or report_type == 'CancerXOME_Unmatched':

        repgen = cxumRepGen(trigger_file, outdir)

    elif report_type == 'Victor CancerXOME Matched' or report_type == 'Victor_CancerXOME_Matched':

        repgen = vcxmRepGen(trigger_file, outdir)

    elif report_type == 'Victor CancerXOME Unmatched' or report_type == 'Victor_CancerXOME_Unmatched':

        repgen = vcxumRepGen(trigger_file, outdir)


    else:
        print("report type '%s' is not supported" % report_type)
        sys.exit(1)

    repgen.generateReport()

if __name__ == "__main__":
    main()

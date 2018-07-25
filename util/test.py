import openpyxl

infile = 'test5.xlsx'
outfile = 'test7.xlsx'

SOMATIC_PEPTIDES_START_ROW = 10

somatic_peptide_records = [['g1','g3','g5'],['g11','g13','g15'],['g21','g23','g25']];


def write_somatic_peptides_sheet():
    sheet = xfile.get_sheet_by_name('Somatic Peptides')
    current_row = SOMATIC_PEPTIDES_START_ROW
    for row in somatic_peptide_records:
        a = 'A' + str(current_row )
        b = 'B' + str(current_row )
        c = 'C' + str(current_row )
        sheet[a] = row[0]
        sheet[b] = row[1]
        sheet[c] = row[2]
        current_row += 1
        
def write_overview_sheet():

    sheet = xfile.get_sheet_by_name('Overview')

    sheet['B14'] = 'waterfall'
    sheet['C14'] = 'pizza'
    sheet['D14'] = 23498249.24892



xfile = openpyxl.load_workbook(infile)

write_somatic_peptides_sheet()

xfile.save(outfile)


## write to somatic mutations sheet







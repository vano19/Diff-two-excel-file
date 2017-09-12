import xlrd, xlwt

# Demonstration of copy2 patch for xlutils 1.4.1

# Context:
# xlutils.copy.copy(xlrd_workbook) -> xlwt_workbook
# copy2(xlrd_workbook) -> (xlwt_workbook, style_list)
# style_list is a conversion of xlrd_workbook.xf_list to xlwt-compatible styles

# Step 1: Create an input file for the demo
##def create_input_file():
##    wtbook = xlwt.Workbook()
##    wtsheet = wtbook.add_sheet(u'First')
##    colours = 'white black red green blue pink turquoise yellow'.split()
##    fancy_styles = [xlwt.easyxf(
##        'font: name Times New Roman, italic on;'
##        'pattern: pattern solid, fore_colour %s;'
##         % colour) for colour in colours]
##    for rowx in range(8):
##        wtsheet.write(rowx, 0, rowx)
##        wtsheet.write(rowx, 1, colours[rowx], fancy_styles[rowx])
##    wtbook.save('demo_copy2_in.xls')
##
# Step 2: Copy the file, changing data content
# ('pink' -> 'MAGENTA', 'turquoise' -> 'CYAN')
# without changing the formatting

from xlutils.filter import process,XLRDReader,XLWTWriter

# Patch: add this function to the end of xlutils/copy.py
def copy2(wb):
    w = XLWTWriter()
    process(
        XLRDReader(wb,'unknown.xls'),
        w
        )
    return w.output[0][1], w.style_list

def update_content():
    rdbook = xlrd.open_workbook('base_.xls', formatting_info=True)
    sheetx = 0
    rdsheet = rdbook.sheet_by_index(sheetx)
    wtbook, style_list = copy2(rdbook)
    wtsheet = wtbook.get_sheet(sheetx)
##    fixups = [(0, 0, 'MAGENTA'), (1, 1, 'CYAN')]
##    for rowx, colx, value in fixups:
##        xf_index = rdsheet.cell_xf_index(rowx, colx)
##        wtsheet.write(rowx, colx, value, style_list[xf_index])
    wtbook.save('demo_copy2_out.xls')

##create_input_file()
update_content()

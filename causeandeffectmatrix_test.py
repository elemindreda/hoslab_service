import xlsxwriter
import pandas
import numpy
import os

spspfilepath = os.getcwd()+'//testing_files//PG2_GC_MP.csv'
rprelayfilepath = os.getcwd()+'//testing_files//PG2_GC_Relay.csv'
syssystemfilepath = os.getcwd()+'//testing_files//PG2_GC_System.csv'

filen = 'Cause_and_Effect_Matrix.xlsx'
workbook = xlsxwriter.Workbook(filen)
ws = workbook.add_worksheet()
data = pandas.read_csv(spspfilepath,sep = ';')
data = data[data.Active != 0]
print(data.values)


##################
### formatting ###
##################

formatheadings = workbook.add_format()
formatheadings.set_bold(True)

border = workbook.add_format()
border.set_border(style=1)
border.set_align('center')

topborder = workbook.add_format()
topborder.set_top(1)

topborderheading = workbook.add_format()
topborderheading.set_top(1)
topborderheading.set_bold(True)

botborder = workbook.add_format()
botborder.set_bottom(1)

normal = workbook.add_format()
normal.set_align('left')

normalc = workbook.add_format()
normalc.set_align('center')

borderhighlightlatch = workbook.add_format()
borderhighlightlatch.set_border(style=1)
borderhighlightlatch.set_align('center')
borderhighlightlatch.set_fg_color('#FF0000')

borderhighlightfault = workbook.add_format()
borderhighlightfault.set_border(style=1)
borderhighlightfault.set_align('center')
borderhighlightfault.set_fg_color('#FFFF00')

borderhighlightAF = workbook.add_format()
borderhighlightAF.set_border(style=1)
borderhighlightAF.set_align('center')
borderhighlightAF.set_fg_color('#0000FF')

ws.set_paper(8)
ws.set_landscape()
ws.set_page_view()
ws.set_margins(top=1.60, left=0.36, right=0.36, bottom=1.0)


workbook.close()

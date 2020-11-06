def create(filen, spspfilepath, rprelayfilepath, syssystemfilepath,):
    import xlsxwriter
    import pandas
    import numpy


    try:
        filen = filen[:len(filen)-5]
        filen = filen + '_Cause_and_Effect_Matrix.xlsx'
        workbook = xlsxwriter.Workbook(filen)
        ws = workbook.add_worksheet()
        data = pandas.read_csv(spspfilepath)
        
    except:
        pass
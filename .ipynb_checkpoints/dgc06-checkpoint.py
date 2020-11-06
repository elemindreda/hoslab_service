def resource_path(relative_path):
    import os
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def activesensors(active):
    for i in range(0, len(active)):
        if active[i] == '1':
            active[i] = i

    active = active[1:]
    active = [e for e in active if e != '0']
    return active


def mp(device, address):
    if device == 'DP':
        mp = 'DP ' + str(address)

    else:
        mp = 'AP ' + str(address)

    return mp


def measurerange(unit, measuringrange):
    mr = '( 0 - ' + str(measuringrange) + unit + ')'
    return mr


def threshold(unit, th, coav):
    if coav == '1':
        coav = 'CV'
    else:
        coav = 'AV'
    thresh = str(th) + ' ' + unit + ', ' + str(coav)
    return thresh


def stage(device, relay):
    stg = device + ' ' + str(relay)
    return stg


def latching(latch1, latch2, latch3, latch4):
    latch = []
    latch = [latch1, latch2, latch3, latch4]
    for i in range(0, len(latch)):
        if latch[i] == '1':
            latch[i] = 'Yes'
            for i in range(0, i):
                latch[i] = 'Yes'

        else:
            latch[i] = 'No'
    return latch


def alarmfalling(AF1, AF2, AF3, AF4):
    AF = []
    AF = [str(AF1), str(AF2), str(AF3), str(AF4)]
    for i in range(0, len(AF)):
        if AF[i] == '1':
            AF[i] = 'Yes'
        else:
            AF[i] = 'No'
    return AF


def fault(fault1, fault2, fault3, fault4):
    fault = []
    fault = [str(fault1), str(fault2), str(fault3), str(fault4)]
    for i in range(0, len(fault)):
        if fault[i] == '1':
            fault[i] = 'Yes'
        else:
            fault[i] = 'No'
    return fault


def relaynumber(SRorPR, number):
    relayident = str(SRorPR) + ' ' + str(number)
    return relayident


def energized(ncno):
    ncno = int(ncno)
    if ncno == 1:
        ncno = 'Energized'
    else:
        ncno = 'De-Energized'
    return ncno


def devicecheck(device, nput):
    if device == 'SR':
        device = 'BI'
    else:
        device = 'DI'
    nPUT = device + ' ' + str(nput)
    return nPUT


## Def funct, comment when testing using sole path.
def internaldoc(filen, spspfilepath, rprelayfilepath, syssystemfilepath, name, projectname, custo, commissioningtech,
                office):
    import csv
    import xlsxwriter
    import shutil

    ################################IMPORTING spSP.csv################################

    ########FILE PATHS UNCOMMENT WHEN PROGRAMMING FROM WORK.
    ##     spspfilepath = os.path.realpath('W:\\google drive\\CSV2EXCEL\\DGC06\\Full System Test_3_8_16GC_SP.csv')
    ##     rprelayfilepath = os.path.realpath('W:\\google drive\\CSV2EXCEL\\DGC06\\Full System Test_3_8_16GC_Relay.csv')
    ##     syssystemfilepath = os.path.realpath('W:\\google drive\\CSV2EXCEL\\DGC06\\Full System Test_3_8_16System.csv')

    ########FILE PATHS UNCOMMENT WHEN PROGRAMMING FROM HOME.
    ##     spspfilepath = os.path.realpath('C:\\Users\\waghu\\google drive\\CSV2EXCEL\\DGC06\\Full System Test_3_8_16GC_SP.csv')
    ##     rprelayfilepath = os.path.realpath('C:\\Users\\waghu\\google drive\\CSV2EXCEL\\DGC06\\Full System Test_3_8_16GC_Relay.csv')
    ##     syssystemfilepath = os.path.realpath('C:\\Users\\waghu\\google drive\\CSV2EXCEL\\DGC06\\Full System Test_3_8_16System.csv')
    ##     workbook = xlsxwriter.Workbook('programming_details.xlsx')
    ########REMINDER LEAVE BOTH WORK & HOME FILE PATHS COMMENTED WHEN COMPILING.

    workbook = xlsxwriter.Workbook(filen)
    ws = workbook.add_worksheet()

    f = open(spspfilepath)
    csv_f = csv.reader(f, delimiter=';')
    raw = []
    active = []
    heading = []
    for row in csv_f:
        raw.append(row)
    raw = [e for e in raw if e]
    column = [[row[i] for row in raw] for i in range(len(raw[0]))]
    active = column[2]
    device = column[0]

    active = activesensors(active)
    sensorstotal = []
    for i in range(0, len(active)):
        sensorstotal.append(raw[(active[i])])

    ##set vars

    ##setting formats
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

    ##page setup values
    ws.set_paper(9)
    ws.set_portrait()
    ws.set_page_view()
    ws.set_margins(top=1.60, left=0.36, right=0.36, bottom=1.0)

    if office == 1:
        header = '&L&G&"Century Gothic,Regular"\nHoslab PTY LTD\n&"Calibri"&18 MSR PG2 \n' \
                 '&"Calibri"&14 Commissioning sheet &R8-10 35 Higginbotham rd \n ' \
                 'Gladesville NSW 2111 \n Australia \n Ph:(02) 9815 3555'
    else:
        header = '&L\n&G&"Century Gothic,Regular"\nHoslab NZ PTY LTD\n(A division Hoslab PTY LTD)&C\n' \
                 '&"Calibri"&18 MSR PG2 \n&"Calibri"&14 Commissioning sheet &R\n ' \
                 'INSERT ROAD HERE \n LOL DOES NZ EVEN HAVE STATES \n New Zealand \n Ph:NZPHONENUMBER'
    footer = '&L www.gasalarm.com.au &C&P &R ' + name + '\n &D'
    logo = resource_path('logo_hoslab.png')
    ws.set_header(header, {'image_left': logo})
    ws.set_footer(footer)

    ##Set Columns for the templatespSP.
    ws.set_column('A:A', 18)
    ws.set_column(1, 21, 3.2)
    ws.set_default_row(15.2)

    ##write heading for templatespSP
    ws.write('A10', 'MP Parameters', formatheadings)

    y = 2  # Vertical Sheet Location variable

    ## CUSTOMER / COMMSSIONING DATA
    ws.write(2, 0, 'Customer:', normal)
    ws.write(3, 0, 'Project:', normal)
    ws.write(4, 0, 'Programmer:', normal)
    ws.write(5, 0, 'Commssioning Tech:', normal)
    ws.merge_range(2, 1, 2, 20, custo, normalc)
    ws.merge_range(3, 1, 3, 20, projectname, normalc)
    ws.merge_range(4, 1, 4, 20, name, normalc)
    ws.merge_range(5, 1, 5, 20, commissioningtech, normalc)
    ##     ws.merge_range(6,1,6,20,building,normalc)
    ## WRITE TO EXCEL FILE MP PARAMETERS
    ## INCLUDES ALL MP + LATCHING + TH + ST + ANALOG OUTPUT INFO.
    for i in range(0, len(active)):
        activeprint = sensorstotal[i]

        MP = mp(activeprint[0], activeprint[1])
        name = activeprint[5]
        measuringrange = measurerange(activeprint[6], activeprint[7])

        TH1 = threshold(activeprint[6], activeprint[8], activeprint[9])
        AF = alarmfalling(activeprint[10], activeprint[13], activeprint[16], activeprint[19])
        latch = latching(str(activeprint[24]), str(activeprint[25]), str(activeprint[26]), str(activeprint[27]))
        FAULT = fault(str(activeprint[28]), str(activeprint[29]), str(activeprint[30]), str(activeprint[31]))
        stage1 = stage(activeprint[32], activeprint[33])
        delayon = 'Alarm Delay On: ' + activeprint[21] + ' sec'
        delayoff = 'Alarm Delay Off: ' + activeprint[21] + ' sec'

        TH2 = threshold(activeprint[6], activeprint[11], activeprint[12])

        stage2 = stage(activeprint[34], activeprint[35])

        TH3 = threshold(activeprint[6], activeprint[14], activeprint[15])

        stage3 = stage(activeprint[36], activeprint[37])

        TH4 = threshold(activeprint[6], activeprint[17], activeprint[18])

        stage4 = stage(activeprint[38], activeprint[39])

        Hysteresis = activeprint[20] + ' ' + activeprint[6]
        AO1 = activeprint[40]
        AO2 = activeprint[41]

        ws.write(y + 10, 0, MP, formatheadings)
        ws.write(y + 11, 0, name, normal)
        ws.write(y + 12, 0, measuringrange, normal)
        ws.merge_range(y + 12, 2, y + 12, 5, 'Thresholds:', normal)
        ws.merge_range(y + 13, 2, y + 13, 5, 'Relays:', normal)
        ws.merge_range(y + 14, 2, y + 14, 5, 'Latching:', normal)
        ws.merge_range(y + 16, 2, y + 16, 5, 'Alarm Falling:', normal)
        ws.merge_range(y + 15, 2, y + 15, 5, 'Activate on fault:', normal)
        ws.merge_range(y + 18, 2, y + 18, 6, 'Analog Output 1:', normal)
        ws.merge_range(y + 19, 2, y + 19, 6, 'Analog Output 2:', normal)
        ws.merge_range(y + 17, 2, y + 17, 6, 'Hysteresis:', normal)
        ##          ws.merge_range(y+12,6,y+12,8,'1',border)
        ##          ws.merge_range(y+12,9,y+12,11,'2',border)
        ##          ws.merge_range(y+12,12,y+12,14,'3',border)
        ##          ws.merge_range(y+12,15,y+12,17,'4',border)
        ws.merge_range(y + 12, 6, y + 12, 8, TH1, border)
        ws.merge_range(y + 12, 9, y + 12, 11, TH2, border)
        ws.merge_range(y + 12, 12, y + 12, 14, TH3, border)
        ws.merge_range(y + 12, 15, y + 12, 17, TH4, border)

        ws.merge_range(y + 13, 6, y + 13, 8, stage1, border)
        ws.merge_range(y + 13, 9, y + 13, 11, stage2, border)
        ws.merge_range(y + 13, 12, y + 13, 14, stage3, border)
        ws.merge_range(y + 13, 15, y + 13, 17, stage4, border)

        for i in range(0, 4):
            if latch[i] == 'Yes':
                ws.merge_range(y + 14, 6 + 3 * i, y + 14, 8 + 3 * i, latch[i], borderhighlightlatch)

            else:
                ws.merge_range(y + 14, 6 + 3 * i, y + 14, 8 + 3 * i, latch[i], border)

        for i in range(0, 4):

            if AF[i] == 'Yes':
                ws.merge_range(y + 16, 6 + 3 * i, y + 16, 8 + 3 * i, AF[i], borderhighlightAF)

            else:
                ws.merge_range(y + 16, 6 + 3 * i, y + 16, 8 + 3 * i, AF[i], border)

        for i in range(0, 4):

            if FAULT[i] == 'Yes':
                ws.merge_range(y + 15, 6 + 3 * i, y + 15, 8 + 3 * i, FAULT[i], borderhighlightfault)

            else:
                ws.merge_range(y + 15, 6 + 3 * i, y + 15, 8 + 3 * i, FAULT[i], border)

        ##          ws.merge_range(y+5,5,y+5,7,SF[0],border)
        ##          ws.merge_range(y+5,8,y+5,10,SF[1],border)
        ##          ws.merge_range(y+5,11,y+5,13,SF[2],border)
        ##          ws.merge_range(y+5,14,y+5,16,SF[3],border)
        ws.merge_range(y + 17, 7, y + 17, 8, Hysteresis, normal)
        ws.merge_range(y + 18, 7, y + 18, 8, AO1, normal)
        ws.merge_range(y + 19, 7, y + 19, 8, AO2, normal)
        ws.merge_range(y + 17, 11, y + 17, 19, delayon, normal)
        ws.merge_range(y + 18, 11, y + 18, 19, delayoff, normal)
        y = y + 11

    ##READ RELAY PARAMETERS
    f = open(rprelayfilepath)
    csv_f = csv.reader(f, delimiter=';')
    raw = []
    active = []
    heading = []
    for row in csv_f:
        raw.append(row)
    raw = [e for e in raw if e]
    column = [[row[i] for row in raw] for i in range(len(raw[0]))]
    active = column[2]
    device = column[0]
    active = activesensors(active)
    relaystotal = []
    for i in range(0, len(active)):
        relaystotal.append(raw[(active[i])])

    y = (44 - y % 44) + y  ## Horizontal Component of iteration
    ## WRITE TO EXCEL FILE RELAY PARAMETERS
    ws.write(y, 0, 'Relay Parameters', formatheadings)
    y = y + 2
    for i in range(0, len(active)):
        activeprint = relaystotal[i]
        ##assigning&formatting relay properties to variables for printing
        staticflash = activeprint[4]
        if staticflash == '1':
            staticflash = 'Flashing'
        else:
            staticflash = 'Static'
        relayident = relaynumber(activeprint[0], activeprint[1])
        ncno = energized(activeprint[3])
        manualondi = devicecheck(activeprint[7], activeprint[8])
        manualoffdi = devicecheck(activeprint[9], activeprint[10])
        delayontime = activeprint[11] + ' sec'
        delayofftime = activeprint[12] + ' sec'
        if activeprint[13] == '1':
            autoresetorrecur = 'Automatic Recurrence'
        else:
            autoresetorrecur = 'Automatic Reset'
        automatictime = activeprint[14] + ' sec'
        automaticresetdi = devicecheck(activeprint[15], activeprint[16])

        ##printing to workbook
        ws.merge_range(y + 2, 1, y + 2, 4, 'Manual On DI: ', border)
        ws.merge_range(y + 3, 1, y + 3, 4, 'Manual Off DI: ', border)
        ws.merge_range(y + 2, 8, y + 2, 11, 'Delay On time: ', border)
        ws.merge_range(y + 3, 8, y + 3, 11, 'Delay Off time: ', border)

        ws.write(y + 2, 0, relayident, topborderheading)
        ws.write(y + 3, 0, ncno, normal)
        ws.write(y + 4, 0, staticflash, normal)
        ws.merge_range(y + 2, 5, y + 2, 6, manualondi, border)
        ws.merge_range(y + 3, 5, y + 3, 6, manualoffdi, border)
        ws.merge_range(y + 2, 12, y + 2, 13, delayontime, border)
        ws.merge_range(y + 3, 12, y + 3, 13, delayofftime, border)
        ws.merge_range(y + 2, 15, y + 2, 19, autoresetorrecur, border)
        ws.merge_range(y + 3, 15, y + 3, 16, 'Time: ', border)
        ws.merge_range(y + 4, 15, y + 4, 16, 'DI:', border)
        ws.merge_range(y + 3, 17, y + 3, 19, automatictime, border)
        ws.merge_range(y + 4, 17, y + 4, 19, automaticresetdi, border)
        ws.write(y + 2, 7, '', topborder)
        ws.write(y + 2, 14, '', topborder)
        ws.write(y + 2, 20, '', topborder)
        y = y + 4

    ##READ SYSTEM PARAMETERS
    f = open(syssystemfilepath)
    csv_f = csv.reader(f, delimiter=';')
    raw = []
    syspara = []
    ao = []
    relaymult = []
    for row in csv_f:
        raw.append(row)
    for i in range(0, 5):
        syspara.append(raw[i])
    column = [[row[i] for row in syspara] for i in range(len(syspara[0]))]

    y = (44 - y % 44) + y
    #### Write to xslx sheet
    sysprops = column[1]
    ws.write(y, 0, 'System Parameters', formatheadings)
    ws.write(y + 1, 0, 'Power On Time:', normal)
    ws.merge_range(y + 1, 1, y + 1, 2, sysprops[0] + ' Sec', normal)

    for i in range(7, 23):
        ao.append(raw[i])
    ao = [[row[i] for row in ao] for i in range(len(ao[0]))]
    AOR = ao[1]
    AOAVCV = ao[2]
    AOFUNC = ao[3]

    for i in range(24, 44):
        relaymult.append(raw[i])
    relaymult = [[row[i] for row in relaymult] for i in range(len(relaymult[0]))]
    rdevin = relaymult[1]
    rin = relaymult[2]
    rdevout = relaymult[3]
    rout = relaymult[4]
    ws.merge_range(y + 4, 1, y + 4, 8, 'Analog Outputs', border)
    ws.merge_range(y + 5, 1, y + 5, 2, 'AO#', border)
    ws.merge_range(y + 5, 3, y + 5, 4, 'Range', border)
    ws.merge_range(y + 5, 5, y + 5, 6, 'AV/CV', border)
    ws.merge_range(y + 5, 7, y + 5, 8, 'Function', border)
    ws.merge_range(y + 4, 11, y + 4, 16, 'Relay Multiplications', border)
    for i in range(1, 17):
        ws.merge_range(y + i + 5, 1, y + i + 5, 2, 'AO' + str(i), border)
        ws.merge_range(y + i + 5, 3, y + i + 5, 4, AOR[i - 1] + '%', border)
        if str(AOAVCV[(i - 1)]) == '1':
            AOAVCV[i - 1] = 'CV'
        else:
            AOAVCV[i - 1] = 'AV'
        ws.merge_range(y + i + 5, 5, y + i + 5, 6, AOAVCV[i - 1], border)
        if str(AOFUNC[i - 1]) == '0':
            AOFUNC[i - 1] = 'Min'
        elif str(AOFUNC[i - 1]) == '1':
            AOFUNC[i - 1] = 'AV'
        else:
            AOFUNC[i - 1] = 'Max'
        ws.merge_range(y + i + 5, 7, y + i + 5, 8, AOFUNC[i - 1], border)
    for i in range(1, 21):
        ws.merge_range(y + i + 4, 11, y + i + 4, 12, str(i), border)
        ws.merge_range(y + i + 4, 13, y + i + 4, 14, rdevin[i - 1] + ' ' + rin[i - 1], border)
        ws.merge_range(y + i + 4, 15, y + i + 4, 16, rdevout[i - 1] + ' ' + rout[i - 1], border)

    workbook.close()
    return






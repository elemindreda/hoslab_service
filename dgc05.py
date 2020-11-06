def resource_path(relative_path):
    import os
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def faultfix(fault, stage, L):
    fault = [int(i) for i in fault]
    stage = [int(i) for i in stage]
    for i in range(0, (len(fault))):
        if fault[i] == 1:
            fault[i] = stage[i]
    fault = fault[0:L]
    return fault


def faultfix2(fault1, fault2, fault3, fault4, fault5, L):
    fault = []
    faultraw = (fault1, fault2, fault3, fault4, fault5)
    faultraw = transpose(faultraw)
    for i in range(0, L):
        fault.append([x for x in faultraw[i] if x != 0])
    for i in range(0, L):
        fault[i] = [str(x) for x in fault[i]]
    for i in range(0, L):
        if not fault[i]:
            fault[i] = 'None'
        else:
            fault[i] = ', '.join(fault[i])
    fault = [e for e in fault if e]
    for i in range(0, L):
        if fault[i] == 'None':
            fault[i] == 'None'
        else:
            fault[i] = 'Relay(s) No. ' + fault[i]
    return (fault)


def internaldoc(filen, spspfilepath, rprelayfilepath, syssystemfilepath, author, projectname, subheading, office):
    import csv
    import xlsxwriter
    import shutil
    import os

    ################################IMPORTING spSP.csv################################
    f = open(spspfilepath)
    csv_f = csv.reader(f, delimiter=';')
    raw = []
    active = []
    heading = []
    for row in csv_f:
        raw.append(row)
    raw = [e for e in raw if e]
    column = transpose(raw)
    active = column[1]
    active = [x for x in active if x != '0']
    L = len(active)
    active = active[2:L]
    ##truncate all junk + format data for printing##
    spaddress = column[0]
    spaddress = spaddress[2:L]

    gastype = column[2]
    gastype = gastype[2:L]

    measuringrange = column[3]
    measuringrange = measuringrange[2:L]

    linear = column[4]
    linear = linear[2:L]
    linear = linearlist(linear)

    threshold1 = column[5]
    threshold1 = threshold1[2:L]

    threshold2 = column[6]
    threshold2 = threshold2[2:L]
    threshold3 = column[7]
    threshold3 = threshold3[2:L]
    threshold4 = column[8]
    threshold4 = threshold4[2:L]
    threshold5 = column[9]
    threshold5 = threshold5[2:L]
    hysteresis = column[10]
    hysteresis = hysteresis[2:L]
    delayontime = column[11]
    delayontime = delayontime[2:L]
    delayofftime = column[12]
    delayofftime = delayofftime[2:L]

    camode = column[13]
    camode = camode[2:L]
    camode = camodelist(camode)

    analogoutput = column[14]
    analogoutput = analogoutput[2:L]
    mp = mpcalc(L)

    # analogoutput = templatewriter.analogoutput(analogoutput)
    stage1 = column[15]
    stage1 = stage1[2:L]
    stage2 = column[16]
    stage2 = stage2[2:L]
    stage3 = column[17]
    stage3 = stage3[2:L]
    stage4 = column[18]
    stage4 = stage4[2:L]
    stage5 = column[19]
    stage5 = stage5[2:L]
    fault1 = column[20]
    fault1 = fault1[2:L]
    fault2 = column[21]
    fault2 = fault2[2:L]
    fault3 = column[22]
    fault3 = fault3[2:L]
    fault4 = column[23]
    fault4 = fault4[2:L]
    fault5 = column[24]
    fault5 = fault5[2:L]
    disable = column[25]
    disable = disable[2:L]
    unit = column[26]
    unit = unit[2:L]
    af = column[27]
    af = af[2:L]
    L = int(len(af))
    af = afcheck(af, L)

    for i in range(0, L):
        if int(threshold1[i]) > int(measuringrange[i]) or int(threshold2[i]) > int(measuringrange[i]) or int(
                threshold3[i]) > int(measuringrange[i]) or int(threshold4[i]) > int(measuringrange[i]) or int(
                threshold5[i]) > int(measuringrange[i]):
            threshold1[i] = int(threshold1[i]) / 10
            threshold2[i] = int(threshold2[i]) / 10
            threshold3[i] = int(threshold3[i]) / 10
            threshold4[i] = int(threshold4[i]) / 10
            threshold5[i] = int(threshold5[i]) / 10
            hysteresis[i] = int(hysteresis[i]) / 10

    measuringrange = measuringrangelist(measuringrange, unit)
    gastype = gastypelist(gastype)

    ################################IMPORTING rpRelay.csv################################
    workingrelays = []
    raw2 = []
    checkrow = []

    relay = []

    workingrelays.extend(stage1)
    workingrelays.extend(stage2)
    workingrelays.extend(stage3)
    workingrelays.extend(stage4)
    workingrelays.extend(stage5)

    workingrelays = cleanuplist(workingrelays)

    L2 = len(workingrelays)
    g = open(rprelayfilepath)
    csv_g = csv.reader(g, delimiter=';')
    for row in csv_g:
        raw2.append(row)
    raw2 = [e for e in raw2 if e]
    column2 = transpose(raw2)
    checkrow = column2[0]
    checkrow = checkrow[2:32]

    col2 = column2[1]
    col2 = col2[2:32]
    col3 = column2[2]
    col3 = col3[2:32]
    col4 = column2[3]
    col4 = col4[2:32]
    col5 = column2[4]
    col5 = col5[2:32]
    col6 = column2[5]
    col6 = col6[2:32]
    col7 = column2[6]
    col7 = col7[2:32]
    col8 = column2[7]
    col8 = col8[2:32]
    col9 = column2[8]
    col9 = col9[2:32]
    col10 = column2[9]
    col10 = col10[2:32]
    col11 = column2[10]
    col11 = col11[2:32]
    col12 = column2[11]
    col12 = col12[2:32]

    for i in range(0, L2):
        for j in range(0, 29):
            if (str(workingrelays[i]) == checkrow[j]):
                relay.append(workingrelays[i])

    relaymode = returnrightlist(col2, workingrelays, checkrow, L2)
    staticflash = returnrightlist(col3, workingrelays, checkrow, L2)
    latchmode = returnrightlist(col4, workingrelays, checkrow, L2)
    autoresettime = returnrightlist(col5, workingrelays, checkrow, L2)
    autorecurrence = returnrightlist(col6, workingrelays, checkrow, L2)
    resetdiginput = returnrightlist(col7, workingrelays, checkrow, L2)
    manualondiginput = returnrightlist(col8, workingrelays, checkrow, L2)
    manualoffdiginput = returnrightlist(col9, workingrelays, checkrow, L2)
    delayontimerp = returnrightlist(col10, workingrelays, checkrow, L2)
    delayofftimerp = returnrightlist(col11, workingrelays, checkrow, L2)
    relaystatus = returnrightlist(col12, workingrelays, checkrow, L2)
    ##     print('workingrelays = ',workingrelays)
    ##     print('relaymode = ',relaymode)
    ##     print('staticflash = ', staticflash)
    ##     print('latchmode = ',latchmode)
    ##     print('autoresettime = ',autoresettime)
    ##     print('autorecurrence = ',autorecurrence)
    ##     print('resetdiginput = ',resetdiginput)
    ##     print('manualondiginput = ',manualondiginput)
    ##     print('manualoffdiginput = ',manualoffdiginput)
    ##     print('delayontimerp = ', delayontimerp)
    ##     print('delayofftimerp = ',delayofftimerp)
    ##     print('relaystatus = ',relaystatus)

    relaymode = relaymodecheck(relaymode, L2)
    latchmode = latchingmodefix(latchmode, L2)
    externalrelay = manualondiginput
    externalrelay = externalrelaycheck(manualondiginput, manualoffdiginput, externalrelay, L2)
    ##     print('externalrelay = ',externalrelay)

    ################################IMPORTING sysSystem.csv################################
    h = open(syssystemfilepath)
    relaymult = []
    relaymultin = []
    relaymultout = []
    epmodules = []
    sysdata = []
    csv_h = csv.reader(h, delimiter=';')
    raw3 = []
    for row in csv_h:
        raw3.append(row)

    for i in range(53, 74):
        relaymult.append(raw3[i])
    relaymult = transpose(relaymult)
    relaymultin = relaymult[1]
    relaymultout = relaymult[2]

    for i in range(13, 37):
        epmodules.append(raw3[i])
    epmodules = transpose(epmodules)
    epmodules = epmodules[1]

    for i in range(0, 11):
        sysdata.append(raw3[i])
    sysdata = transpose(sysdata)
    sysdata = sysdata[1]

    mp = mpcalc(L)
    mp2 = mp2calc(L2)

    ################################ Writing spSP.csv to file ################################
    workbook = xlsxwriter.Workbook(filen)
    ws = workbook.add_worksheet()
    # set vars

    # setting formats
    formatheadings = workbook.add_format()
    formatheadings.set_bold(True)
    border = workbook.add_format()
    border.set_border(style=1)
    border.set_align('center')

    # page setup values
    ws.set_paper(9)
    ws.set_portrait()
    ws.set_page_view()
    ws.set_margins(top=1.77, left=0.36, right=0.36, bottom=0.75)
    if office == 1:
        header = '&L&G&"Century Gothic,Regular"\nHoslab\n&"Calibri"&18MSR PG2 \n' \
                 '&"Calibri"&14Commissioning sheet &R8-10 35 Higginbotham rd \n ' \
                 'Gladesville NSW 2111 \n Australia \n Ph:(02) 9815 3555'
    else:
        header = '&L\n&G&"Century Gothic,Regular"\nHoslab\n(A division of Hoslab)&C\n' \
                 '&"Calibri"&18 MSR PG2 \n&"Calibri"&14 Commissioning sheet &R\n ' \
                 'PO Box 825 Ryde \n NSW 1680 \n Australia \n Ph: +61 (02) 9816 3555'  
    footer = '&L www.hoslab.com.au &C&P &R ' + author + '\n &D'
    logo = resource_path('logo_gasalarm.png')
    ws.set_header(header, {'image_left': logo})
    ws.set_footer(footer)

    # Set Columns for the templatespSP.
    ws.set_column('A:A', 18)
    ws.set_column(1, 21, 3.2)

    # write heading for templatespSP
    ws.write('A1', 'MP Parameters', formatheadings)

    # Template loops + nested
    i = 0
    j = 0
    x = 1
    z = 0
    c1 = 0
    c2 = 0
    d = 0
    last = 21
    if (L < 4):
        if (L % 4 == 1):
            last = 6
        elif (L % 4 == 2):
            last = 11
        elif (L % 4 == 3):
            last = 16
    ##     print(af)
    ##     print (L)
    y = 0
    while (z < (mp)):
        for j in range(x, x + 17):
            if (j == x + 13) or (j == x + 14):
                continue
            else:
                for i in range(1, last, 5):
                    if (j < 18):
                        c1 = int((i - 1) / 5)
                    elif (c1 == (L)):
                        break
                    else:
                        c1 = int((i - 1) / 5) + d
                    if (j == 1) or (((j + 25) % 45) == 0) or ((j - 1) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 4), spaddress[c1], border)
                    elif (j == 2) or (((j + 24) % 45) == 0) or ((j - 2) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 4), gastype[c1], border)
                    elif (j == 3) or (((j + 23) % 45) == 0) or ((j - 3) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 4), measuringrange[c1], border)
                    elif (j == 4) or (((j + 22) % 45) == 0) or ((j - 4) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 4), linear[c1], border)
                    elif (j == 5) or (((j + 21) % 45) == 0) or ((j - 5) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 4), threshold1[c1], border)
                    elif (j == 6) or (((j + 20) % 45) == 0) or ((j - 6) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 4), threshold2[c1], border)
                    elif (j == 7) or (((j + 19) % 45) == 0) or ((j - 7) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 4), threshold3[c1], border)
                    elif (j == 8) or (((j + 18) % 45) == 0) or ((j - 8) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 4), threshold4[c1], border)
                    elif (j == 9) or (((j + 17) % 45) == 0) or ((j - 9) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 4), threshold5[c1], border)
                    elif (j == 10) or (((j + 16) % 45) == 0) or ((j - 10) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 4), hysteresis[c1], border)
                    elif (j == 11) or (((j + 15) % 45) == 0) or ((j - 11) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 4), delayontime[c1], border)
                    elif (j == 12) or (((j + 14) % 45) == 0) or ((j - 12) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 4), delayofftime[c1], border)
                    elif (j == 13) or (((j + 13) % 45) == 0) or ((j - 13) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 4), camode[c1], border)
                    elif (j == 16) or (((j + 10) % 45) == 0) or ((j - 16) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 4), analogoutput[c1], border)
                    elif (j == 17) or (((j + 9) % 45) == 0) or ((j - 17) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 4), af[c1], border)
        for j in range(x + 13, x + 16):
            for i in range(1, last, 5):
                if (j < 18):
                    c2 = int((i - 1) / 5)
                else:
                    c2 = int((i - 1) / 5) + d
                if (j == 14) or ((j + 12) % 45 == 0) or ((j - 14) % 45 == 0):
                    ws.write(j, i, fault1[c2], border)
                if (j == 15) or ((j + 11) % 45 == 0) or ((j - 15) % 45 == 0):
                    ws.write(j, i, stage1[c2], border)
            for i in range(2, last, 5):
                if (j < 18):
                    c2 = int((i - 1) / 5)
                else:
                    c2 = int((i - 1) / 5) + d
                if (j == 14) or ((j + 12) % 45 == 0) or ((j - 14) % 45 == 0):
                    ws.write(j, i, fault2[c2], border)
                if (j == 15) or ((j + 11) % 45 == 0) or ((j - 15) % 45 == 0):
                    ws.write(j, i, stage2[c2], border)
            for i in range(3, last, 5):
                if (j < 18):
                    c2 = int((i - 1) / 5)
                else:
                    c2 = int((i - 1) / 5) + d
                if (j == 14) or ((j + 12) % 45 == 0) or ((j - 14) % 45 == 0):
                    ws.write(j, i, fault3[c2], border)
                if (j == 15) or ((j + 11) % 45 == 0) or ((j - 15) % 45 == 0):
                    ws.write(j, i, stage3[c2], border)
            for i in range(4, last, 5):
                if (j < 18):
                    c2 = int((i - 1) / 5)
                else:
                    c2 = int((i - 1) / 5) + d
                if (j == 14) or ((j + 12) % 45 == 0) or ((j - 14) % 45 == 0):
                    ws.write(j, i, fault4[c2], border)
                if (j == 15) or ((j + 11) % 45 == 0) or ((j - 15) % 45 == 0):
                    ws.write(j, i, stage4[c2], border)
            for i in range(5, last, 5):
                if (j < 18):
                    c2 = int((i - 1) / 5)
                else:
                    c2 = int((i - 1) / 5) + d
                if (j == 14) or ((j + 12) % 45 == 0) or ((j - 14) % 45 == 0):
                    ws.write(j, i, fault5[c2], border)
                if (j == 15) or ((j + 11) % 45 == 0) or ((j - 15) % 45 == 0):
                    ws.write(j, i, stage5[c2], border)
        c2 = 0
        if (c1 == (L - 2)) or (c1 == L - 3) or (c1 == L - 4):
            d = d + 4
            if (L % 4 == 1):
                last = 6
            elif (L % 4 == 2):
                last = 11
            elif (L % 4 == 3):
                last = 16
        else:
            d = d + 4

        # write titles/heading rows with border formattings
        ws.write(x, 0, 'MP Number', border)
        ws.write(x + 1, 0, 'Gas Type', border)
        ws.write(x + 2, 0, 'Measuring Range', border)
        ws.write(x + 3, 0, 'Signal', border)
        ws.write(x + 4, 0, 'Threshold 1', border)
        ws.write(x + 5, 0, 'Threshold 2', border)
        ws.write(x + 6, 0, 'Threshold 3', border)
        ws.write(x + 7, 0, 'Threshold 4', border)
        ws.write(x + 8, 0, 'Threshold 5', border)
        ws.write(x + 9, 0, 'Hysteresis', border)
        ws.write(x + 10, 0, 'Delay ON', border)
        ws.write(x + 11, 0, 'Delay OFF', border)
        ws.write(x + 12, 0, 'Mode', border)
        ws.write(x + 13, 0, 'Alarm(1,2,3,4,5)/Fault', border)
        ws.write(x + 14, 0, 'A1 A2 A3 A4 A5', border)
        ws.write(x + 15, 0, 'Analog Output', border)
        ws.write(x + 16, 0, 'Alarm Falling', border)
        y = y + 1
        if (y % 2 == 1):
            x = x + 19
        else:
            x = x + 26
        z = z + 1

    ############################### Writing rpRelay.csv to file ################################

    # Static Heading#
    if ((x - 1) % 45 == 0):
        ws.write(x - 1, 0, 'Relay Parameters', formatheadings)

    else:
        ws.write(x + 2, 0, 'Relay Parameters', formatheadings)
        x = x + 3
    # Template loops + nested arguments
    d = 0
    last = 19
    c1 = 0
    c2 = 0
    z = 0
    stator = x

    if (L2 < 6) or (L2 == 6):
        if (L2 % 6 == 1):
            last = 4
        elif (L2 % 6 == 2):
            last = 7
        elif (L2 % 6 == 3):
            last = 10
        elif (L2 % 6 == 4):
            last = 13
        elif (L2 % 6 == 5):
            last = 16
        elif (L2 % 6 == 0):
            last = 19
    print('MP2 = ', mp2)
    print('L2 = ', L2)
    print('relay = ', relay)
    while (z < (mp2)):
        for j in range(x, x + 8):
            ##            if (x-1)%45==0:
            ##                 t=1
            if (j == x + 4):
                continue
            else:
                for i in range(1, last, 3):
                    if ((j - stator) < 9):
                        c1 = int((i - 1) / 3)
                    else:
                        c1 = int(((i - 1) / 3) + d)
                    if (c1 == L2):
                        break
                    if ((j - 1) % 45 == 0) or ((j - 12) % 45 == 0) or ((j - 23) % 45 == 0) or ((j - 34) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 2), relay[c1], border)
                    elif ((j - 2) % 45 == 0) or ((j - 13) % 45 == 0) or ((j - 24) % 45 == 0) or ((j - 35) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 2), relaymode[c1], border)
                    elif ((j - 3) % 45 == 0) or ((j - 14) % 45 == 0) or ((j - 25) % 45 == 0) or ((j - 36) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 2), staticflash[c1], border)
                    elif ((j - 4) % 45 == 0) or ((j - 15) % 45 == 0) or ((j - 26) % 45 == 0) or ((j - 37) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 2), latchmode[c1], border)
                    elif ((j - 6) % 45 == 0) or ((j - 17) % 45 == 0) or ((j - 28) % 45 == 0) or ((j - 39) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 2), externalrelay[c1], border)
                    elif ((j - 7) % 45 == 0) or ((j - 18) % 45 == 0) or ((j - 29) % 45 == 0) or ((j - 40) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 2), delayontimerp[c1], border)
                    elif ((j - 8) % 45 == 0) or ((j - 19) % 45 == 0) or ((j - 30) % 45 == 0) or ((j - 41) % 45 == 0):
                        ws.merge_range(j, i, j, (i + 2), delayofftimerp[c1], border)
            for i in range(1, last, 3):
                if ((j - stator) < 9):
                    c2 = int((i - 1) / 3)
                else:
                    c2 = int((i - 1) / 3) + d

                if (c2 == L2):
                    break
                ws.write(x + 4, i, autoresettime[c2], border)
                ws.write(x + 4, i + 1, autorecurrence[c2], border)
                ws.write(x + 4, i + 2, resetdiginput[c2], border)

        if (c1 == (L2 - 1)) or (c1 == (L2 - 2)) or (c1 == (L2 - 3)) or (c1 == (L2 - 4)) or (c1 == (L2 - 5)) or (
                c1 == (L2 - 6)):
            d = d + 6
            if (L2 % 6 == 1):
                last = 4
            elif (L2 % 6 == 2):
                last = 7
            elif (L2 % 6 == 3):
                last = 10
            elif (L2 % 6 == 4):
                last = 13
            elif (L2 % 6 == 5):
                last = 16
        else:
            d = d + 6

        # write titles and row headings.
        ws.write(x, 0, 'Relay Number', border)
        ws.write(x + 1, 0, 'Relay Mode', border)
        ws.write(x + 2, 0, 'Static Flash', border)
        ws.write(x + 3, 0, 'Latching Mode', border)
        ws.write(x + 4, 0, 'Time-Recurr.-DI', border)
        ws.write(x + 5, 0, 'External Relay', border)
        ws.write(x + 6, 0, 'Delay ON', border)
        ws.write(x + 7, 0, 'Delay OFF', border)
        z = z + 1
        pspacerem = pspaceremaining(x)
        if (pspacerem < 13):
            x = x + pspacerem + 1
        else:
            x = x + 11

    ############################### Writing sysSystem.csv to file ################################

    z = 1
    pspacerem = pspaceremaining(x)
    if (pspacerem < 26):
        x = x + pspacerem
    # left column headings + Title
    ws.write(x + 1, 0, 'System Parameters', formatheadings)
    ws.write(x + 2, 0, 'Service Mode', border)
    ws.write(x + 3, 0, 'Software Version', border)
    ws.write(x + 4, 0, 'Next Maint. Date', border)
    ws.write(x + 5, 0, 'Service Phone', border)
    ws.write(x + 6, 0, 'AV-Overlay (time)', border)
    ws.write(x + 7, 0, 'AV-Overlay (ppm)', border)
    ws.write(x + 8, 0, 'AV Time', border)
    ws.write(x + 9, 0, 'Timesystem', border)
    ws.write(x + 10, 0, 'Time', border)
    ws.write(x + 11, 0, 'Date', border)
    ws.write(x + 12, 0, 'Customers Pass', border)
    ws.write(x + 13, 0, 'Failure Relay', border)
    ws.write(x + 14, 0, 'Power On Time', border)
    # merge cells
    ws.merge_range(x + 2, 1, x + 2, 4, 'Off', border)
    ws.merge_range(x + 3, 1, x + 3, 4, 'DGC05_9_02C', border)

    while z < 12:
        if (sysdata[z - 1] == '0'):
            sysdata[z - 1] = 'EU'
        elif (sysdata[z - 1] == '1'):
            sysdata[z - 1] = 'US'
        ws.merge_range(z + x + 3, 1, z + x + 3, 4, sysdata[z - 1], border)
        z = z + 1
    # merge + format EP module headings + cells
    ws.merge_range(x + 1, 6, x + 1, 9, 'EP Modules', formatheadings)
    z = 0
    while z < 24:
        y = str(z)
        ws.merge_range(z + x + 2, 6, z + x + 2, 9, 'EP module ' + y + '', border)
        if epmodules[z] == '0':
            epmodules[z] = 'Not active'
        elif epmodules[z] == '1':
            epmodules[z] = 'Relay active'
        elif epmodules[z] == '2':
            epmodules[z] = 'Rel. + MP'
        ws.merge_range(z + x + 2, 10, z + x + 2, 12, epmodules[z], border)
        z = z + 1
    # merge + format Relay multiplication tables
    z = 0
    ws.merge_range(x + 1, 14, x + 1, 18, 'Relay Mult.', formatheadings)
    ws.write(x + 1, 19, 'IN', formatheadings)
    ws.write(x + 1, 20, 'OUT', formatheadings)
    while z < 20:
        ws.merge_range(z + x + 2, 14, z + x + 2, 18, 'Relay No. 0-30', border)
        ws.write(z + x + 2, 19, relaymultin[z + 1], border)
        ws.write(z + x + 2, 20, relaymultout[z + 1], border)
        z = z + 1
    workbook.close()


def customerdoc(filen, spspfilepath, rprelayfilepath, syssystemfilepath, author, projectname, subheading, office):
    import csv
    import xlsxwriter
    import shutil
    import os

    ################################IMPORTING spSP.csv################################
    f = open(spspfilepath)
    csv_f = csv.reader(f, delimiter=';')
    raw = []
    active = []
    heading = []
    for row in csv_f:
        raw.append(row)
    raw = [e for e in raw if e]
    column = transpose(raw)
    active = column[1]
    active = [x for x in active if x != '0']
    L = len(active)
    active = active[2:L]

    ##truncate all junk + format data for printing##
    spaddress = column[0]
    spaddress = spaddress[2:L]
    gastype = column[2]
    gastype = gastype[2:L]

    measuringrange = column[3]
    measuringrange = measuringrange[2:L]
    linear = column[4]
    linear = linear[2:L]
    linear = linearlist(linear)
    threshold1 = column[5]
    threshold1 = threshold1[2:L]
    threshold2 = column[6]
    threshold2 = threshold2[2:L]
    threshold3 = column[7]
    threshold3 = threshold3[2:L]
    threshold4 = column[8]
    threshold4 = threshold4[2:L]
    threshold5 = column[9]
    threshold5 = threshold5[2:L]
    hysteresis = column[10]
    hysteresis = hysteresis[2:L]
    delayontime = column[11]
    delayontime = delayontime[2:L]
    delayofftime = column[12]
    delayofftime = delayofftime[2:L]
    camode = column[13]
    camode = camode[2:L]
    camode = camodelist(camode)
    analogoutput = column[14]
    analogoutput = analogoutput[2:L]
    stage1 = column[15]
    stage1 = stage1[2:L]
    stage2 = column[16]
    stage2 = stage2[2:L]
    stage3 = column[17]
    stage3 = stage3[2:L]
    stage4 = column[18]
    stage4 = stage4[2:L]
    stage5 = column[19]
    stage5 = stage5[2:L]
    disable = column[25]
    disable = disable[2:L]
    unit = column[26]
    unit = unit[2:L]
    af = column[27]
    af = af[2:L]

    L = int(len(af))
    fault1 = column[20]
    fault1 = fault1[2:100]
    fault2 = column[21]
    fault2 = fault2[2:100]
    fault3 = column[22]
    fault3 = fault3[2:100]
    fault4 = column[23]
    fault4 = fault4[2:100]
    fault5 = column[24]
    fault5 = fault5[2:100]
    stage1 = column[15]
    stage1 = stage1[2:100]
    stage2 = column[16]
    stage2 = stage2[2:100]
    stage3 = column[17]
    stage3 = stage3[2:100]
    stage4 = column[18]
    stage4 = stage4[2:100]
    stage5 = column[19]
    stage5 = stage5[2:100]

    fault1 = faultfix(fault1, stage1, L)
    fault2 = faultfix(fault2, stage2, L)
    fault3 = faultfix(fault3, stage3, L)
    fault4 = faultfix(fault4, stage4, L)
    fault5 = faultfix(fault5, stage5, L)
    fault = faultfix2(fault1, fault2, fault3, fault4, fault5, L)
    af = afcheck(af, L)
    for i in range(0, L):
        if int(threshold1[i]) > int(measuringrange[i]) or int(threshold2[i]) > int(measuringrange[i]) or int(
                threshold3[i]) > int(measuringrange[i]) or int(threshold4[i]) > int(measuringrange[i]) or int(
                threshold5[i]) > int(measuringrange[i]):
            threshold1[i] = int(threshold1[i]) / 10
            threshold2[i] = int(threshold2[i]) / 10
            threshold3[i] = int(threshold3[i]) / 10
            threshold4[i] = int(threshold4[i]) / 10
            threshold5[i] = int(threshold5[i]) / 10
            hysteresis[i] = int(hysteresis[i]) / 10
    measuringrange = measuringrangelist(measuringrange, unit)
    gastype = gastypelist(gastype)
    mp = mpcalc(L)

    ################################ Writing spSP.csv to file ################################
    workbook = xlsxwriter.Workbook(filen)
    ws = workbook.add_worksheet()
    # set vars

    # setting formats
    formatheadings = workbook.add_format()
    formatheadings.set_bold(True)

    format1 = workbook.add_format()
    format1.set_bg_color('#e6e6e6')
    format1.set_align('center')
    format1.set_border(style=1)
    format2 = workbook.add_format()
    format2.set_bg_color('#FFFFFF')
    format2.set_align('center')
    format2.set_border(style=1)

    # page setup values
    ws.set_paper(9)
    ws.set_portrait()
    ws.set_page_view()
    ws.set_margins(top=1.77, left=0.36, right=0.36, bottom=0.75)
    if office == 1:
        header = '&L\n          &G&"Century Gothic,Regular"\n             GasAlarm Systems\n(A division of ALVI Technologies)&C\n&"Calibri"&16' + projectname + '\n&12' + subheading + '&R\n2/79 Station Road \n Seven Hills NSW 2147 \n Australia \n Ph:(02) 9838 7220'
    else:
        header = '&L\n          &G&"Century Gothic,Regular"\n             GasAlarm Systems SA\n(A division of ALVI Technologies)&C\n&"Calibri"&16' + projectname + '\n&12' + subheading + '&R\n12 Bideford ave \n Clarence Gardens SA 5039 \n Australia \n Ph: 0416202261'
    footer = '&L www.gasalarm.com.au &C&P &R ' + author + '\n &D'
    logo = resource_path('logo_gasalarm.png')
    ws.set_header(header, {'image_left': logo})
    ws.set_footer(footer)

    # Set Columns for the templatespSP.
    ws.set_column('A:A', 18)
    ws.set_column(1, 21, 3.2)

    # write heading for templatespSP
    ws.write('A1', 'MP Parameters', formatheadings)
    # Template loops + nested
    i = 0
    j = 0
    x = 1
    z = 0
    c1 = 0
    c2 = 0
    d = 0
    last = 21
    if (L < 4):
        if (L % 4 == 1):
            last = 6
        elif (L % 4 == 2):
            last = 11
        elif (L % 4 == 3):
            last = 16
    y = 0
    while (z < (mp)):
        for j in range(x, x + 21):
            for i in range(1, last, 5):
                if (j < 18):
                    c1 = int((i - 1) / 5)
                elif (c1 == (L)):
                    break
                else:
                    c1 = int((i - 1) / 5) + d
                # print('c = ',c)
                if (j == 1) or ((j + 21) % 45 == 0) or ((j - 1) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), spaddress[c1], format2)
                elif (j == 2) or ((j + 20) % 45 == 0) or ((j - 2) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), gastype[c1], format1)
                elif (j == 3) or ((j + 19) % 45 == 0) or ((j - 3) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), measuringrange[c1], format2)
                elif (j == 4) or ((j + 18) % 45 == 0) or ((j - 4) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), linear[c1], format1)
                elif (j == 5) or ((j + 17) % 45 == 0) or ((j - 5) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), threshold1[c1], format2)
                elif (j == 6) or ((j + 16) % 45 == 0) or ((j - 6) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), threshold2[c1], format1)
                elif (j == 7) or ((j + 15) % 45 == 0) or ((j - 7) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), threshold3[c1], format2)
                elif (j == 8) or ((j + 14) % 45 == 0) or ((j - 8) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), threshold4[c1], format1)
                elif (j == 9) or ((j + 13) % 45 == 0) or ((j - 9) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), threshold5[c1], format2)
                elif (j == 10) or ((j + 12) % 45 == 0) or ((j - 10) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), hysteresis[c1], format1)
                elif (j == 11) or ((j + 11) % 45 == 0) or ((j - 11) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), delayontime[c1], format2)
                elif (j == 12) or ((j + 10) % 45 == 0) or ((j - 12) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), delayofftime[c1], format1)
                elif (j == 13) or ((j + 9) % 45 == 0) or ((j - 13) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), camode[c1], format2)
                elif (j == 14) or ((j + 8) % 45 == 0) or ((j - 14) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), fault[c1], format1)
                elif (j == 15) or ((j + 7) % 45 == 0) or ((j - 15) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), stage1[c1], format2)
                elif (j == 16) or ((j + 6) % 45 == 0) or ((j - 16) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), stage2[c1], format1)
                elif (j == 17) or ((j + 5) % 45 == 0) or ((j - 17) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), stage3[c1], format2)
                elif (j == 18) or ((j + 4) % 45 == 0) or ((j - 18) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), stage4[c1], format1)
                elif (j == 19) or ((j + 3) % 45 == 0) or ((j - 19) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), stage5[c1], format2)
                elif (j == 20) or ((j + 2) % 45 == 0) or ((j - 20) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), analogoutput[c1], format1)
                elif (j == 21) or ((j + 1) % 45 == 0) or ((j - 21) % 45 == 0):
                    ws.merge_range(j, i, j, (i + 4), af[c1], format2)
        if (c1 == (L - 2)) or (c1 == (L - 3)) or (c1 == (L - 4)):
            d = d + 4
            if (L % 4 == 1):
                last = 6
            elif (L % 4 == 2):
                last = 11
            elif (L % 4 == 3):
                last = 16
        else:
            d = d + 4

        # write titles/heading rows with border formattings
        ws.write(x, 0, 'Sensor Number', format2)
        ws.write(x + 1, 0, 'Gas Type', format1)
        ws.write(x + 2, 0, 'Measuring Range', format2)
        ws.write(x + 3, 0, 'Sensor Signal', format1)
        ws.write(x + 4, 0, 'Alarm Threshold 1', format2)
        ws.write(x + 5, 0, 'Alarm Threshold 2', format1)
        ws.write(x + 6, 0, 'Alarm Threshold 3', format2)
        ws.write(x + 7, 0, 'Alarm Threshold 4', format1)
        ws.write(x + 8, 0, 'Alarm Threshold 5', format2)
        ws.write(x + 9, 0, 'Hysteresis', format1)
        ws.write(x + 10, 0, 'Delay ON time (sec)', format2)
        ws.write(x + 11, 0, 'Delay OFF time (sec)', format1)
        ws.write(x + 12, 0, 'Control Mode', format2)
        ws.write(x + 13, 0, 'Sensor Fault', format1)
        ws.write(x + 14, 0, 'Alarm 1: Relay', format2)
        ws.write(x + 15, 0, 'Alarm 2: Relay', format1)
        ws.write(x + 16, 0, 'Alarm 3: Relay', format2)
        ws.write(x + 17, 0, 'Alarm 4: Relay', format1)
        ws.write(x + 18, 0, 'Alarm 5: Relay', format2)
        ws.write(x + 19, 0, 'Analog Output', format1)
        ws.write(x + 20, 0, 'Alarm Falling', format2)
        y = y + 1
        if (y % 2 == 1):
            x = x + 23
        else:
            x = x + 22
        z = z + 1
    workbook.close()


def afcheck(af, L):
    for i in range(0, L):
        if af[i] == '0':
            af[i] = 'None'
        elif af[i] == '1':
            af[i] = 'Threshold: 1'
        elif af[i] == '2':
            af[i] = 'Threshold: 1, 2'
        elif af[i] == '3':
            af[i] = 'Threshold: 1, 2, 3'
        elif af[i] == '4':
            af[i] = 'Threshold: 1, 2, 3, 4'
        elif af[i] == '5':
            af[i] = 'Threshold: 1, 2, 3, 4, 5'
    return af


def externalrelaycheck(manualondiginput, manualoffdiginput, externalrelay, L2):
    for i in range(0, L2):
        externalrelay[i] = 'On = ' + manualondiginput[i] + ', Off = ' + manualoffdiginput[i]
    return externalrelay


def relaymodecheck(relaymode, L2):
    for i in range(0, L2):
        if relaymode[i] == '0':
            relaymode[i] = 'De-energised'
        else:
            relaymode[i] = 'Energised'
    return relaymode


def latchingmodefix(latchingmode, L2):
    for i in range(0, L2):
        if latchingmode[i] == '0':
            latchingmode[i] = 'No'
        else:
            latchingmode[i] = 'Yes'
    return latchingmode


def transpose(matrix):
    if not matrix: return []
    return [[row[i] for row in matrix] for i in range(len(matrix[0]))]


def mpcalc(L):
    if (L % 4 == 0):
        mp = int(L / 4)
    elif (L % 4 == 1):
        mp = int(L / 4) + 0.75
    elif (L % 4 == 2):
        mp = int(L / 4) + 0.5
    elif (L % 4 == 3):
        mp = int(L / 4) + 0.25
    return mp


def mp2calc(L):
    if (L % 6 == 0):
        mp2 = int(L / 6)
    elif (L % 6 == 1):
        mp2 = int((L / 6) + (5 / 6))
    elif (L % 6 == 2):
        mp2 = int((L / 6) + (4 / 6))
    elif (L % 6 == 3):
        mp2 = int((L / 6) + (3 / 6))
    elif (L % 6 == 4):
        mp2 = int((L / 6) + (2 / 6))
    elif (L % 6 == 5):
        mp2 = int((L / 6) + (1 / 6))
    return mp2


def pspaceremaining(x):
    pspaceremaining = 45 - (x % 45)
    if pspaceremaining == 45:
        pspaceremaining = 0
    return pspaceremaining


def gastypelist(gastype):
    L = len(gastype)
    for i in range(0, L):
        if (gastype[i] == '0'):
            gastype[i] = 'CO (Toxic)'
        elif (gastype[i] == '1'):
            gastype[i] = 'EX (Explosive)'
        elif (gastype[i] == '2'):
            gastype[i] = 'NO (Toxic)'
        elif gastype[i] == '3':
            gastype[i] = 'NO2 (Toxic)'
        elif gastype[i] == '4':
            gastype[i] = 'NH3 (Toxic)'
        elif gastype[i] == '5':
            gastype[i] = 'O2 (Toxic)'
        elif gastype[i] == '6':
            gastype[i] = 'CO2'
        elif gastype[i] == '7':
            gastype[i] = 'SO2'
        elif gastype[i] == '8':
            gastype[i] = 'H2S'
        elif gastype[i] == '9':
            gastype[i] = 'CL2'
        elif gastype[i] == '10':
            gastype[i] = 'ETO'
        elif gastype[i] == '11':
            gastype[i] = 'VOC'
        elif gastype[i] == '12':
            gastype[i] = 'R4XX'
        elif gastype[i] == '13':
            gastype[i] = 'R5XX'
        elif gastype[i] == '14':
            gastype[i] = 'R11'
        elif gastype[i] == '15':
            gastype[i] = 'R123'
        elif gastype[i] == '16':
            gastype[i] = 'R134'
        elif gastype[i] == '17':
            gastype[i] = 'R22'
        elif gastype[i] == '18':
            gastype[i] = 'TEM'
        elif gastype[i] == '19':
            gastype[i] = 'RH'
        elif gastype[i] == '20':
            gastype[i] = 'TOX'
        elif gastype[i] == '21':
            gastype[i] = 'CH4'
        elif gastype[i] == '22':
            gastype[i] = 'VAP'
        elif gastype[i] == '23':
            gastype[i] = 'EXIR (explo. Infra Red)'
        elif gastype[i] == '24':
            gastype[i] = 'NF3'
        elif gastype[i] == '25':
            gastype[i] = 'PCT'
        elif gastype[i] == '26':
            gastype[i] = 'NO GAS'
    return gastype


def measuringrangelist(measuringrange, unit):
    L = len(measuringrange)
    for i in range(0, L):
        if unit[i] == '0':
            unit[i] = 'ppm'
        elif unit[i] == '1':
            unit[i] = '%LEL'
        elif unit[i] == '2':
            unit[i] = 'VOL%'
        elif unit[i] == '3':
            unit[i] = 'deg (F)'
        elif unit[i] == '4':
            unit[i] = '%RH'
        elif unit[i] == '5':
            unit[i] = '%'
        elif unit[i] == '6':
            unit[i] = 'ppk'
        elif unit[i] == '7':
            unit[i] = 'deg(C)'
    for i in range(0, L):
        measuringrange[i] = '0 to ' + measuringrange[i] + ' ' + unit[i]
    return measuringrange


def linearlist(linear):
    L = len(linear)
    for i in range(0, L):
        if linear[i] == '1':
            linear[i] = 'Linear'
        else:
            linear[i] = 'Non Linear'
    return linear


def camodelist(camode):
    L = len(camode)
    for i in range(0, L):
        if camode[i] == '1':
            camode[i] = 'Current Value'
        else:
            camode[i] = 'Average Value'
    return camode


def analogoutputlist(analogoutput):
    L = len(analogoutput)
    for i in range(0, L):
        if analogoutput[i] == '1':
            # might need to change to 4-20 / 2-10 V for the printed value.
            analogoutput[i] = 'True'
        else:
            analogoutput[i] = 'False'
    return analogoutput


def test(lista, L):
    for i in range(9, L + 1):
        lista.append(i)
    return lista


def cleanuplist(workingrelays):
    newlist = []
    for i in workingrelays:
        if i not in newlist:
            newlist.append(i)
    newlist = [int(i) for i in newlist]
    for i in newlist:
        if i == 0:
            newlist.remove(0)
    newlist.sort()

    return newlist


def returnrightlist(column, workingrelays, checkrow, L2):
    formattedcolumn = []
    workingrelaysint = [int(i) for i in workingrelays]
    for i in range(0, L2):
        for j in range(0, 29):
            if (str(workingrelays[i]) == checkrow[j]):
                formattedcolumn.append(column[(workingrelaysint[i] - 1)])
    return formattedcolumn

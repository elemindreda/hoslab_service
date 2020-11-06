from tkinter import *
from tkinter import messagebox
import tkinter.filedialog as filedialog
import os
import dgc05
import dgc06
import causeandeffectmatrix
import datetime

##'ConfigParser for first time setup'
##import ConfigParser
mygui = Tk()
spsp = StringVar()
rprelay = StringVar()
syssystem = StringVar()
global spspfilepath
global rprelayfilepath
global syssystemfilepath
spspfilepath = ''
rprelayfilepath = ''
syssystemfilepath = ''
internal = IntVar()
customer = IntVar()
dgcselect = IntVar()
officeselect = StringVar()
c_and_e = IntVar()

mygui.title('CSV TO EXCEL v2.026, ( ͡° ͜ʖ ͡°)')
mygui.geometry('400x700+700+250')


def truncate(filepath):
    if (len(filepath) > 209):
        base = os.path.basename(filepath)
        fdir = os.path.dirname(filepath)
        x = 209 - (len(base) + 3)
        filepath = fdir[:x] + '...' + base
    return filepath


def spspfile():
    b = spsp.get()
    global spspfilepath
    spspfilepath = filedialog.askopenfilename(filetypes=(('CSV Files', '.csv'), ('All Files', '*,*')))
    fdir = truncate(spspfilepath)
    spsplabel = Label(spsptitle, text=fdir, font=("calibri", 8), wraplength=250).grid(row=3, column=0)
    spspfilepath = os.path.realpath(spspfilepath)
    return spspfilepath


def rprelayfile():
    b = rprelay.get()
    global rprelayfilepath
    rprelayfilepath = filedialog.askopenfilename(filetypes=(('CSV Files', '.csv'), ('All Files', '*,*')))
    fdir = truncate(rprelayfilepath)
    rprelaylabel = Label(rprelaytitle, text=fdir, font=("calibri", 8), wraplength=250).grid(row=6, column=0)
    return rprelayfilepath


def syssystemfile():
    b = syssystem.get()
    global syssystemfilepath
    syssystemfilepath = filedialog.askopenfilename(filetypes=(('CSV Files', '.csv'), ('All Files', '*,*')))
    fdir = truncate(syssystemfilepath)
    syssystemlabel = Label(syssystemtitle, text=fdir, font=("calibri", 8), wraplength=250).grid(row=9, column=0)
    return syssystemfilepath


def shutdownprogram():
    quittermessage = messagebox.askyesno(title='Quit', message='Are you sure you want to Quit?')
    if quittermessage == 1:
        mygui.destroy()


'testing code'


##def testgc06():
##    global spspfilepath
##    global rprelayfilepath
##    global syssystemfilepath
##    author = 'sw'
##    projectname='TESTTESTTEST'
##    optionalheading='TEST'
##    spspfilepath = 'sp.csv'
##    rprelayfilepath = 'rp.csv'
##    syssystemfilepath = 'sys.csv'
##    now = datetime.datetime.now()
##    filen = 'Test'+str(now.day)+'-'+str(now.month)+'.xlsx'
##    dgc06.internaldoc(filen,spspfilepath,rprelayfilepath,syssystemfilepath,author,projectname,optionalheading)
##    mygui.destroy()
def create():
    global spspfilepath
    global rprelayfilepath
    global syssystemfilepath
    name = author.get()
    name = str(name)
    projectname = str(proj.get())
    custo = str(cust.get())
    commissioningtech = str(commissiontech.get())
    creator = name
    optionalheading = custo

    office = officeselect.get()
    office = int(office)
    dgc = dgcselect.get()
    dgc = int(dgc)
    cb1 = internal.get()
    cb2 = customer.get()
    causeeffect = c_and_e.get()
    if not author:
        noname = messagebox.showinfo('No name entered', 'enter name before clicking create')
        return
    spspfilen = os.path.basename(spspfilepath)
    spspfilen = spspfilen[(len(spspfilen) - 6):len(spspfilen)]
    spspfilen = spspfilen.lower()
    if spspfilen != 'sp.csv':
        if spspfilen != 'mp.csv':
            wrongfilesp = messagebox.showinfo('Wrong file selected', 'please ensure correct file is selected, sp.csv')
            return
    rprelayfilen = os.path.basename(rprelayfilepath)
    rprelayfilen = rprelayfilen[(len(rprelayfilen) - 9):len(rprelayfilen)]
    rprelayfilen = rprelayfilen.lower()
    if rprelayfilen != 'relay.csv':
        wrongfilerp = messagebox.showinfo('Wrong file selected', 'please ensure correct file is selected, relay.csv')
        return
    syssystemfilen = os.path.basename(syssystemfilepath)
    syssystemfilen = syssystemfilen[(len(syssystemfilen) - 10):len(syssystemfilen)]
    syssystemfilen = syssystemfilen.lower()
    if syssystemfilen != 'system.csv':
        wrongfilesys = messagebox.showinfo('Wrong file selected', 'please ensure correct file is selected, system.csv')
        return
    if dgc == 1:
        if cb1 == 1:
            filen = filedialog.asksaveasfilename(defaultextension='.xlsx', title='Save Internal Document')
            dgc05.internaldoc(filen, spspfilepath, rprelayfilepath, syssystemfilepath, creator, projectname,
                              optionalheading, office)
        if cb2 == 1:
            filen1 = filedialog.asksaveasfilename(defaultextension='.xlsx', title='Save Customer Document')
            dgc05.customerdoc(filen1, spspfilepath, rprelayfilepath, syssystemfilepath, creator, projectname,
                              optionalheading, office)
        return
    if dgc == 2:
        filen = filedialog.asksaveasfilename(defaultextension='.xlsx', title='Save Internal Document')
        dgc06.internaldoc(filen, spspfilepath, rprelayfilepath, syssystemfilepath, name, projectname, custo,
                          commissioningtech, office)
        if causeeffect == 1:
            causeandeffectmatrix.create(filen, spspfilepath, rprelayfilepath, syssystemfilepath,)
        return


def disablecust():
    if dgcselect.get() == 2:
        CUSTOMER_CHECK.config(state='disable')
    elif dgcselect.get() == 1:
        CUSTOMER_CHECK.config(state='normal')
    return


##config = configparser.ConfigParser()


mymenu = Menu(mygui)
filelist = Menu(mygui)
filelist.add_command(label='Quit', command=shutdownprogram)

mymenu.add_cascade(label='File', menu=filelist)
dgcselect.set('2')
Radiobutton(mygui, text="DGC 05", variable=dgcselect, value=1, command=disablecust).grid(row=0, column=0)
Radiobutton(mygui, text="DGC 06", variable=dgcselect, value=2, command=disablecust).grid(row=0, column=1)

INTERNAL_CHECK = Checkbutton(mygui, text="Internal", variable=internal)
INTERNAL_CHECK.grid(row=1, column=0)
CUSTOMER_CHECK = Checkbutton(mygui, text="Customer", variable=customer)
CUSTOMER_CHECK.grid(row=1, column=1)

spsptitle = LabelFrame(mygui, text='Select spsp file name', width=250, height=100).grid(row=3, column=0, pady=(0, 20))
spspbrowse = Button(mygui, text='Browse', command=spspfile).grid(row=3, column=1)

rprelaytitle = LabelFrame(mygui, text='Select rprelay file name', width=250, height=100).grid(row=6, column=0,
                                                                                              pady=(0, 20))
rprelaybrowse = Button(mygui, text='Browse', command=rprelayfile).grid(row=6, column=1)

syssystemtitle = LabelFrame(mygui, text='Select sysSystem.csv file', width=250, height=100).grid(row=9, column=0,
                                                                                                 pady=(0, 20))
syssytembrowse = Button(mygui, text='Browse', command=syssystemfile).grid(row=9, column=1)

writenamehere = Label(mygui, text='Enter programmer name: ', font=12, wraplength=250).grid(row=12, column=0)
author = Entry(mygui)
author.grid(row=12, column=1)
author.focus_set()

custo = Label(mygui, text='Customer: ', font=12, wraplength=250).grid(row=13, column=0)
cust = Entry(mygui)
cust.grid(row=13, column=1)

project = Label(mygui, text='Project: ', font=12, wraplength=250).grid(row=14, column=0)
proj = Entry(mygui)
proj.grid(row=14, column=1)

commissioningtech = Label(mygui, text='Commissioning Tech: ', font=12, wraplength=250).grid(row=15, column=0)
commissiontech = Entry(mygui)
commissiontech.grid(row=15, column=1)

officeselect.set('1')
Radiobutton(mygui, text="Sydney Address", variable=officeselect, value='1').grid(row=16, column=0, sticky=W)
Radiobutton(mygui, text="New Zealand Address", variable=officeselect, value='2').grid(row=16, column=1, sticky=W)

causeandeffectcheck = Checkbutton(mygui, text='Generate Cause and Effect Matrix', variable = c_and_e)
causeandeffectcheck.grid(row = 17, column = 0, pady=(20,20))

create = Button(mygui, text='Create File', command=create).grid(row=18, column=0, columnspan=2, pady=(20, 0))


'testbutton enable'
##test=Button(mygui,text='test',command=testgc06).grid(row=19,column=0)


mygui.config(menu=mymenu)
mygui.mainloop()




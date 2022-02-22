import tkinter
from tkinter import ttk as GUI
from tkinter import E,W
from tkinter import filedialog

def tempBrowse():
    filename = filedialog.askdirectory(initialdir='/',title='Select Folder')
    tempDir.set(filename)
    makeOptionWin()

def outBrowse():
    filename = filedialog.askdirectory(initialdir='/',title='Select Folder')
    outDir.set(filename)
    makeOptionWin()


def makeOptionWin(*args):
    optWin = tkinter.Toplevel(appWin)
    optFrame = GUI.Frame(optWin)
    optFrame.grid(column=0,row=0)

    tempDirLabel = GUI.Label(optFrame,text='Template File Directory:')
    tempDirLabel.grid(column=1,row=1,sticky=W)
    tempDirLabel.grid_configure(padx=5,pady=1)

    
    tempDirEntry = GUI.Entry(optFrame, width=70, textvariable=tempDir)
    tempDirEntry.grid(column=1,row=2,columnspan=10)
    tempDirEntry.grid_configure(padx=5,pady=5)

    tempDirBut = GUI.Button(optFrame, text='...',width=3, command=tempBrowse)
    tempDirBut.grid(column=11,row=2,sticky=W)
    tempDirBut.grid_configure(padx=3,pady=7)

    outDirLabel = GUI.Label(optFrame,text='File Output Directory:')
    outDirLabel.grid(column=1,row=3,sticky=W)
    outDirLabel.grid_configure(padx=5,pady=1)

    
    outDirEntry = GUI.Entry(optFrame, width=70, textvariable=outDir)
    outDirEntry.grid(column=1,row=4,columnspan=10)
    outDirEntry.grid_configure(padx=5,pady=5)

    outDirBut = GUI.Button(optFrame, text='...',width=3,command=outBrowse)
    outDirBut.grid(column=11,row=4,sticky=W)
    outDirBut.grid_configure(padx=3,pady=7)

    enterButton = GUI.Button(optFrame, text='Enter', command=exit)
    enterButton.grid(column=10, row=5,sticky=E)
    enterButton.grid_configure(padx=3,pady=10)

    exitButton = GUI.Button(optFrame, text='Cancel', command=exit)
    exitButton.grid(column=11,row=5,sticky=W)
    exitButton.grid_configure(padx=3,pady=10)


    optWin.mainloop()

def makeMainWin() -> tkinter.Tk:
    Win = tkinter.Tk()
    Win.title('Template Maker v2.0')
    #winIcon = tkinter.PhotoImage(file="C:/Users/ctroc/Documents/Python/TemplateMaker/DANDM.png",format='png')
    #Win.iconphoto(True,winIcon)
    Win.columnconfigure(0,weight=1)
    Win.rowconfigure(0,weight=1)

    winFrame = GUI.Frame(Win)
    winFrame.grid(column=0,row=0)

    input = tkinter.StringVar()
    userEntry = GUI.Entry(winFrame, width=20,textvariable=input)
    userEntry.grid(column=2, row= 1,columnspan=2,sticky=W)
    userEntry.grid_configure(padx=5,pady=10)

    entryLabel = GUI.Label(winFrame,text='Enter Application #:')
    entryLabel.grid(column=1,row=1)
    entryLabel.grid_configure(padx=5,pady=10)

    optionButton = GUI.Button(winFrame, text='Options...', command=makeOptionWin)
    optionButton.grid(column=1, row=2, sticky=W)
    optionButton.grid_configure(padx=3,pady=10)

    enterButton = GUI.Button(winFrame, text='Enter', command=exit)
    enterButton.grid(column=2, row=2)
    enterButton.grid_configure(padx=3,pady=10)

    exitButton = GUI.Button(winFrame, text='Cancel', command=exit)
    exitButton.grid(column=3,row=2)
    exitButton.grid_configure(padx=3,pady=10)
    userEntry.focus()
    return Win

appWin = makeMainWin()
appWin.bind("<Return>",exit)
tempDir = tkinter.StringVar(value='Desktop')
outDir = tkinter.StringVar(value='Desktop')
appWin.mainloop()
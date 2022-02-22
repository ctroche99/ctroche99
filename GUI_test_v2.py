import pathlib
import tkinter
import tkinter.messagebox
from tkinter import ttk as GUI
from tkinter import E,W
from tkinter import filedialog

def initDirectories(FileName: str) -> pathlib.Path: #check if file exists and if not, returns file path to newly made file
    curDir = str(pathlib.Path(__file__).parent.resolve())
    FilePath = curDir+FileName
    FilePath = pathlib.Path(FilePath)
    if FilePath.exists() == False:
        dirOut_file = open(str(FilePath),'w')
        dirOut_file.write(curDir+'\n')
        dirOut_file.write(curDir+'\n')
        dirOut_file.close()
    return FilePath  

class mainWin:
    def __init__(self,root: tkinter.Tk,tempDirectory :str, outDirectory :str,dirPath :pathlib.Path) -> None:
        self.tempDir = tkinter.StringVar(value=tempDirectory)
        self.outDir = tkinter.StringVar(value=outDirectory)
        self.dirPath = dirPath

        self.root = root
        self.root.title('Template Maker v2.1')
        self.root.columnconfigure(0,weight=1)
        self.root.rowconfigure(0,weight=1)

        self.frame = GUI.Frame(self.root)
        self.frame.grid(column=0,row=0)

        self.input = tkinter.StringVar()
        self.entry = GUI.Entry(self.frame, width=20,textvariable=self.input)
        self.entry.grid(column=2,row=1,columnspan=2,sticky=W)
        self.entry.grid_configure(padx=5,pady=5)

        self.entryLabel = GUI.Label(self.frame,text='Enter Application #:')
        self.entryLabel.grid(column=1,row=1)
        self.entryLabel.grid_configure(padx=5,pady=10)

        self.optButn = GUI.Button(self.frame, text='Options...', command=self.makeOptWin)
        self.optButn.grid(column=1, row=2, sticky=W)
        self.optButn.grid_configure(padx=3,pady=10)

        self.entButn = GUI.Button(self.frame, text='Enter', command=self.makeTemp)
        self.entButn.grid(column=2, row=2)
        self.entButn.grid_configure(padx=3,pady=10)

        self.exButn = GUI.Button(self.frame, text='Cancel', command=exit)
        self.exButn.grid(column=3,row=2)
        self.exButn.grid_configure(padx=3,pady=10)
        self.entry.focus()
        self.root.bind("<Return>",self.makeTemp)

    def makeOptWin(self):
        self.optWin = tkinter.Toplevel(self.root)
        self.optApp = optWin(self.root,self.optWin,self.tempDir,self.outDir,self.dirPath)
    
    def makeTemp(self):
        pass

class optWin:
    def __init__(self,master :tkinter.Tk, root :tkinter.Tk,tempDir :tkinter.StringVar,outDir :tkinter.StringVar,dirPath: pathlib.Path):
        
        self.root = root#Option Window Root
        self.frame = GUI.Frame(self.root)
        self.frame.grid(column=0,row=0)
        self.root.transient(master)#Original Window's Root

        self.tempDirLabel = GUI.Label(self.frame,text='Template File Directory:')
        self.tempDirLabel.grid(column=1,row=1,sticky=W)
        self.tempDirLabel.grid_configure(padx=5,pady=1)
    
        self.tempDirEntry = GUI.Entry(self.frame, width=70, textvariable=tempDir)
        self.tempDirEntry.grid(column=1,row=2,columnspan=10)
        self.tempDirEntry.grid_configure(padx=5,pady=5)

        self.tempDirBut = GUI.Button(self.frame, text='...',width=3, command=lambda: self.tempBrowse(tempDir))
        self.tempDirBut.grid(column=11,row=2,sticky=W)
        self.tempDirBut.grid_configure(padx=3,pady=7)

        self.outDirLabel = GUI.Label(self.frame,text='File Output Directory:')
        self.outDirLabel.grid(column=1,row=3,sticky=W)
        self.outDirLabel.grid_configure(padx=5,pady=1)
    
        self.outDirEntry = GUI.Entry(self.frame, width=70, textvariable=outDir)
        self.outDirEntry.grid(column=1,row=4,columnspan=10)
        self.outDirEntry.grid_configure(padx=5,pady=5)

        self.outDirBut = GUI.Button(self.frame, text='...',width=3,command=lambda: self.outBrowse(outDir))
        self.outDirBut.grid(column=11,row=4,sticky=W)
        self.outDirBut.grid_configure(padx=3,pady=7)

        self.entButn = GUI.Button(self.frame, text='Enter', command=lambda: self.entPress(tempDir,outDir,dirPath))
        self.entButn.grid(column=10, row=5,sticky=E)
        self.entButn.grid_configure(padx=3,pady=10)

        self.exButn = GUI.Button(self.frame, text='Cancel', command=self.root.destroy)
        self.exButn.grid(column=11,row=5,sticky=W)
        self.exButn.grid_configure(padx=3,pady=10)

    def tempBrowse(self,directory: tkinter.StringVar):
        filename = filedialog.askdirectory(initialdir='/',title='Select Folder')
        directory.set(filename)

    def outBrowse(self, directory: tkinter.StringVar):
        filename = filedialog.askdirectory(initialdir='/',title='Select Folder')
        directory.set(filename)

    def entPress(self,tempDir :tkinter.StringVar,outDir :tkinter.StringVar,dirPath :pathlib.Path):
        dirOut_file = open(str(dirPath),'w') #write the newly added directories to default file
        dirOut_file.write(tempDir.get()+'\n')
        dirOut_file.write(outDir.get()+'\n')
        dirOut_file.close()
        self.root.destroy()

directoryFilePath = initDirectories('\TemplateMaker_Directories.txt')
dirIn_file = open(str(directoryFilePath),'r')
templateDirectory = dirIn_file.readline()
outputDirectory = dirIn_file.readline()
dirIn_file.close()

root = tkinter.Tk()
mainWindow = mainWin(root,templateDirectory,outputDirectory,directoryFilePath)
root.mainloop()
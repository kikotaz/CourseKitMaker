from tkinter import *
from tkinter import filedialog
from tkinter import font
from tkinter import ttk
from pathlib import Path, PureWindowsPath
from shutil import copyfile
import os
import tkinter as tk
import WordHandler
import re
import ctypes
import threading



firstSubList = ['Assessments','Class Roll','Course Outline','Course Result Summary',
                'Lecture Material','Other Documents','SpreadSheet']
assesmentsSubList = []
assessmentSecondSubList = ['Drafts', 'Moderation Materials', 'Submissions']
assessmentThirdSubList = ['Moderation forms', 'Three Samples']

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def fileOpen() :
    openFileName = filedialog.askopenfilename(
            filetypes=(("Word file", "*.doc"),("All Files","*.*")),
            title = "Choose a file."
            )
    if openFileName:
        filePath.set(openFileName)
        window.mainloop()

def folderCreate(directory):
    try:
        if not(os.path.exists(directory)):
            os.makedirs(directory)
    except OSError:
            errorBox = ctypes.windll.user32.MessageBoxW
            errorBox(None, 'Failed to create directory!!!!!', 0)

def createSub(rootDirectory, folderList):
    try:
        for i in folderList:
            folderName = rootDirectory+"\\"+i
            if not(os.path.exists(folderName)):
                os.makedirs(folderName)
    except OSError:
            errorBox = ctypes.windll.user32.MessageBoxW
            errorBox(None, 'Failed to create sub directory!!!!!', 0)

def createWeek1to12(rootDirectory):
    try:
        for i in range(1, 13):
            folderName = rootDirectory+"\\week"+str(i)
            if not(os.path.exists(folderName)):
                os.makedirs(folderName)
    except OSError:
            errorBox = ctypes.windll.user32.MessageBoxW
            errorBox(None, 'Failed to create week directory!!!!!', 0)

def createAssementsSecondSub(rootDirectory):
    try:
        for i in assessmentSecondSubList:
            folderName = rootDirectory+"\\"+i
            if not(os.path.exists(folderName)):
                os.makedirs(folderName)
    except OSError:
            errorBox = ctypes.windll.user32.MessageBoxW
            errorBox(None, 'Failed to create second sub directory!!!!!', 0)

def isBlank(myString):
    if myString and myString.strip():
        #myString is not None AND myString is not empty or blank
        return False
    #myString is None OR myString is empty or blank
    return True

def createAllFolder():
    print("filePath == " + str(filePath.get()))
    if isBlank(filePath.get()):
        failBox = ctypes.windll.user32.MessageBoxW
        failBox(None, 'The course decriptor file does not selected', 'Message', 0)
    else:
        saveFolderName = filedialog.askdirectory()
        print(saveFolderName)
        if not isBlank(saveFolderName):
            x = threading.Thread(target=doCreateAllFolder, args=(saveFolderName,))
            x.start()
            btnOpenSource.configure(state=tk.DISABLED)
            run_animation()

def doCreateAllFolder(saveFolderName):
    purePath = str(PureWindowsPath(filePath.get()))
    word = WordHandler.WordHandler()
    word.checkCourseDescriptor(purePath.replace('\\', '\\\\'))
    extractList = word.extractData(purePath.replace('\\', '\\\\'))
    for i in range(2, len(extractList)):
        print(extractList[i])
        assesmentsSubList.append(removechars(extractList[i]))
        
    rootFolderName = "{code}-{semester}-{year}"
    rootFolderName = rootFolderName.replace('{code}', removechars(extractList[0]))
    rootFolderName = rootFolderName.replace("{semester}", RadioVariety_1.get())
    rootFolderName = rootFolderName.replace("{year}", comboYear.get())

    rootFolderName = str(PureWindowsPath(saveFolderName)) + "\\" + rootFolderName
    print(rootFolderName)

    folderCreate(rootFolderName)

    #create first sub folder
    createSub(rootFolderName, firstSubList)

    #create assesment sub folder
    createSub(rootFolderName + "\\" + firstSubList[0], assesmentsSubList)

    #create assesment second sub folder
    for i in assesmentsSubList:
        subFolderName = rootFolderName + "\\" + firstSubList[0] + "\\" + i
        createAssementsSecondSub(subFolderName)

        for k in assessmentSecondSubList:
            subSecondFolderName = subFolderName + "\\" + k
            if (k==assessmentSecondSubList[1]):
                #create assesment third sub folder
                createSub(subSecondFolderName + "\\", assessmentThirdSubList)

    #create lecturer meterials sub folder
    createWeek1to12(rootFolderName + "\\" + firstSubList[4])     

    outlineFolder = os.path.abspath(rootFolderName) + '\\Course Outline'
    #create course outline file
    retVal = word.createOutline(removechars(extractList[0]), removechars(extractList[1]), RadioVariety_1.get(),
     comboYear.get(), outlineFolder)
    
    path,filename=os.path.split(filePath.get())
    print("filename == " + filename)
    copyfile(purePath, outlineFolder + "\\" + filename)

    if retVal=="S":
        stop_animation()
        successBox = ctypes.windll.user32.MessageBoxW
        successBox(None, 'The course has been published', 'Message', 0)
    else:
        stop_animation()
        failBox = ctypes.windll.user32.MessageBoxW
        failBox(None, 'The course has been failed', 'Message', 0)

    btnOpenSource.configure(state=tk.NORMAL)

def createOptions():
    frame_0 = tk.Frame(window, background="white")
    frame_0.pack()

    frame_1 = tk.Frame(frame_0)
    frame_1.pack(expand=True, side=LEFT, fill='both', padx=50)

    labelframe = tk.LabelFrame(frame_1, text="Semester")
    labelframe.pack()

    RadioVariety_1.set("none")

    radio1 = tk.Radiobutton(labelframe, text="Semester 1", value="S1", variable=RadioVariety_1)
    radio1.pack()
    radio2 = tk.Radiobutton(labelframe, text="Semester 2", value="S2", variable=RadioVariety_1)
    radio2.pack()
    radio3 = tk.Radiobutton(labelframe, text="Semester 3", value="S3", variable=RadioVariety_1)
    radio3.pack()
    radio1.select()

    frame_2 = tk.Frame(frame_0)
    frame_2.pack(expand=True, side=LEFT, fill='both', padx=50)

    labelframe2 = tk.LabelFrame(frame_2, text="Year")
    labelframe2.pack(expand=True, fill='both')

    combo = ttk.Combobox(labelframe2, width=10, textvariable=comboYear)
    for i in range(2019, 2051):
        combo['values'] = (*combo['values'], i)
        combo.current(0)
    combo.pack()

def removechars(cellvalue):
    text = re.sub(r"[\r\n\t\x07\x0b]", "", cellvalue)
    return text

def run_animation():
    maxFrame = 30
    frames = [PhotoImage(file=resource_path("loading.gif"),format = 'gif -index %i' %(i)) for i in range(maxFrame)]
    def update(ind):
        if ind == maxFrame:
            ind = 0
        frame = frames[ind]
        ind += 1
        labelLoading.configure(image=frame)
        window.after(100, update, ind)
    global canvLoading
    canvLoading=tk.Canvas(window, width=200,height=100, bg="white", bd=0, highlightthickness=0)
    canvLoading.place(x = 150, y = 350)

    labelLoading = Label(canvLoading)
    labelLoading.config(bg='white')
    labelLoading.place(x=65, y=15)

    window.after(0, update, 0)
    window.mainloop()

def stop_animation():
    canvLoading.place_forget()

window = tk.Tk()
window.title("CourseKitMaker")
window.geometry("500x500")
window.configure(bg="white")
window.resizable(FALSE, FALSE)

RadioVariety_1 = StringVar()
comboYear = StringVar()
filePath = StringVar()
fileCourseOutlinePath = StringVar()

imgLogo=tk.PhotoImage(file=resource_path("logo.png"))
label=tk.Label(window, image=imgLogo, borderwidth=0, highlightthickness=0)
label.config(justify=CENTER, pady=150)
label.pack()

LabelsFont = font.Font(family='Time New Roman', size=20, weight='bold')
lblProgName = tk.Label(window, wraplength = 1000, font=LabelsFont, fg="grey", 
                bg="white", text="Course Kit Generator",borderwidth=0, 
                compound="center",highlightthickness=0)
lblProgName.config(justify=CENTER, pady=20)
lblProgName.pack()

imgOpenSource = PhotoImage(file=resource_path("btn_open_file.png"))
btnOpenSource = tk.Button(None, text = "button", image = imgOpenSource, 
                command = fileOpen, borderwidth=0,highlightthickness=0)
btnOpenSource.config(justify=CENTER, pady=20)
btnOpenSource.pack()


LabelsFont = font.Font(family='Time New Roman', size=10, weight='bold')
lblFileName = tk.Label(window, wraplength = 1000, font=LabelsFont, fg="grey",
                bg="white", textvariable=filePath, borderwidth=0,compound="center",highlightthickness=0)
lblFileName.config(justify=CENTER, pady=10)
lblFileName.pack()

createOptions()

lblEmpty = tk.Label(window, wraplength = 1000, font=LabelsFont, fg="grey",
                bg="white", borderwidth=0,compound="center",highlightthickness=0)
lblEmpty.config(justify=CENTER, pady=10)
lblEmpty.pack()

imgGenerate = PhotoImage(file=resource_path("btn_generate.png"))
btnGenerate = tk.Button(None, text = "button", image = imgGenerate,
                command = createAllFolder, borderwidth=0,highlightthickness=0, pady=30)
btnGenerate.config(justify=CENTER, pady=50)
btnGenerate.pack()

canv=tk.Canvas(window, width=800,height=50, bg="blue", bd=0, highlightthickness=0)
canv.place(x = 0, y = 450)
canv.create_rectangle(0, 0, 1001, 50, fill="blue", outline="")

LabelsFont = font.Font(family='Time New Roman', size=12, weight='bold')
lblDeveloper = tk.Label(window, wraplength = 800, font=LabelsFont, fg="white",
                bg="blue", text="@Developed by X1-S3-2019",
                borderwidth=0,compound="center",highlightthickness=0)
lblDeveloper.place(x=120,y=465)

window.mainloop()
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
import datetime
import shutil

firstSubList = ['Assessments','Class Roll','Course Outline','Course Result Summary',
                'Lecture Material','Other Documents','SpreadSheet']
assesmentsSubList = []
assessmentSecondSubList = ['Drafts', 'Moderation Materials', 'Submissions']
assessmentThirdSubList = ['Moderation forms', 'Three Samples']
semesterAndYear = []

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
        fileStatus.set(False)
        filePath.set(openFileName)
        purePath = str(PureWindowsPath(filePath.get()))

        x = threading.Thread(target=doCheckFileFormat, args=(purePath,))
        x.start()

        btnOpenSource.configure(state=tk.DISABLED)
        btnGenerate.configure(state=tk.DISABLED)
        removeButtonEvent()
        run_animation("Checking Format...")
        window.mainloop()

def doCheckFileFormat(purePath):
    word = WordHandler.WordHandler()
    fileStatus.set(word.checkCourseDescriptor(purePath.replace('\\', '\\\\')))
    stop_animation()
    btnOpenSource.configure(state=tk.NORMAL)
    btnGenerate.configure(state=tk.NORMAL)

    createButtonEvent()

def folderCreate(directory):
    try:
        if not(os.path.exists(directory)):
            os.makedirs(directory)
        else:
            shutil.rmtree(directory)
            #os.removedirs(directory)
    except OSError as e:
        print("folderCreate Exception == " + e)
        errorBox = ctypes.windll.user32.MessageBoxW
        errorBox(None, 'Failed to create directory!!!!!', 'Message', 0)

def createSub(rootDirectory, folderList):
    try:
        for i in folderList:
            folderName = rootDirectory+"\\"+i
            if not(os.path.exists(folderName)):
                os.makedirs(folderName)
    except OSError as e:
        print("folderCreate Exception == " + e)
        errorBox = ctypes.windll.user32.MessageBoxW
        errorBox(None, 'Failed to create sub directory!!!!!', 'Message', 0)

def createWeek1to12(rootDirectory):
    try:
        for i in range(1, 13):
            folderName = rootDirectory+"\\week"+str(i)
            if not(os.path.exists(folderName)):
                os.makedirs(folderName)
    except OSError as e:
        print("folderCreate Exception == " + e)
        errorBox = ctypes.windll.user32.MessageBoxW
        errorBox(None, 'Failed to create week directory!!!!!', 'Message', 0)

def createAssementsSecondSub(rootDirectory):
    try:
        for i in assessmentSecondSubList:
            folderName = rootDirectory+"\\"+i
            if not(os.path.exists(folderName)):
                os.makedirs(folderName)
    except OSError:
            errorBox = ctypes.windll.user32.MessageBoxW
            errorBox(None, 'Failed to create second sub directory!!!!!', 'Message', 0)

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
        failBox(None, 'The Course Descriptor file is not uploaded', 'Message', 0)
    else:
        if(fileStatus.get() == True):
            saveFolderName = filedialog.askdirectory()
            print(saveFolderName)
            if not isBlank(saveFolderName):
                x = threading.Thread(target=doCreateAllFolder, args=(saveFolderName,))
                x.start()

                btnOpenSource.configure(state=tk.DISABLED)
                btnGenerate.configure(state=tk.DISABLED)
                removeButtonEvent()
                run_animation("In Progress...")
        else:
            failBox = ctypes.windll.user32.MessageBoxW
            failBox(None, 'The Course Descriptor file format is not correct', 'Message', 0)

def doCreateAllFolder(saveFolderName):
    assesmentsSubList.clear()
    purePath = str(PureWindowsPath(filePath.get()))
    word = WordHandler.WordHandler()
    word.checkCourseDescriptor(purePath.replace('\\', '\\\\'))
    extractList = word.extractData(purePath.replace('\\', '\\\\'))
    for i in range(2, len(extractList)):
        print("extractList == " + extractList[i])
        assesmentsSubList.append(removechars(extractList[i]))
        
    tempData = comboSemesterYear.get()
    splitData = tempData.split(', ')
    selSemester = splitData[0].replace('Semester ', 'S')
    selYear = splitData[1]

    rootFolderName = "{code}-{semester}-{year}"
    rootFolderName = rootFolderName.replace('{code}', removechars(extractList[0]))
    rootFolderName = rootFolderName.replace("{semester}", selSemester)
    rootFolderName = rootFolderName.replace("{year}", selYear)

    rootFolderName = str(PureWindowsPath(saveFolderName)) + "\\" + rootFolderName
    print(rootFolderName)

    folderCreate(rootFolderName)

    #create first sub folder
    createSub(rootFolderName, firstSubList)

    #create assesment sub folder
    createSub(rootFolderName + "\\" + firstSubList[0], assesmentsSubList)

    #create assesment second sub folder
    for i in assesmentsSubList:
        if(isBlank(i)):
            print("Assessment Title is empty.")
        else:
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
    retVal = word.createOutline(removechars(extractList[0]), removechars(extractList[1]), selSemester,
     selYear, outlineFolder)
    
    path,filename=os.path.split(filePath.get())
    print("filename == " + filename)
    copyfile(purePath, outlineFolder + "\\" + filename)

    if retVal=="S":
        stop_animation()
        filePath.set("")
        successBox = ctypes.windll.user32.MessageBoxW
        successBox(None, 'The Course Kit has been published', 'Message', 0)
    else:
        stop_animation()
        failBox = ctypes.windll.user32.MessageBoxW
        failBox(None, 'The Course Kit has failed to publish', 'Message', 0)

    btnOpenSource.configure(state=tk.NORMAL)
    btnGenerate.configure(state=tk.NORMAL)

    createButtonEvent()

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

def run_animation(showingText):
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
    canvLoading=tk.Canvas(window, width=200,height=110, bg="white", bd=0, highlightthickness=0)
    canvLoading.place(x = 150, y = 370)

    labelLoading = Label(canvLoading)
    labelLoading.config(bg='white')
    labelLoading.place(x=65, y=15)

    LabelsFont = font.Font(family='Time New Roman', size=10, weight='bold')
    lblProgressName = tk.Label(canvLoading, wraplength = 1000, font=LabelsFont, fg="red", 
                bg="white", text=showingText,borderwidth=0, 
                compound="center",highlightthickness=0)
    lblProgressName.config(pady=10)
    if(showingText == "In Progress..."):
        lblProgressName.place(x=60, y=80)
    else:
        lblProgressName.place(x=40, y=80)

    window.after(0, update, 0)
    window.mainloop()

def stop_animation():
    lblProgressName.place_forget()
    canvLoading.place_forget()

def callback_motion(event, imgPath, btn):
    imgOpenSourceOver = PhotoImage(file=resource_path(imgPath))
    btn.config(image=imgOpenSourceOver)
    btn.image = imgOpenSourceOver

def callback_leave(event, imgPath, btn):
    imgOpenSourceOver = PhotoImage(file=resource_path(imgPath))
    btn.config(image=imgOpenSourceOver)
    btn.image = imgOpenSourceOver    

def createOptions2():
    frame_0 = tk.Frame(window, background="white")
    frame_0.pack()
    
    frame_2 = tk.Frame(frame_0)
    frame_2.pack(expand=True, side=LEFT, fill='both', padx=50, pady=10)

    labelframe2 = tk.LabelFrame(frame_2, text="Semester and Year")
    labelframe2.pack(expand=True, fill='both')

    currentDate = datetime.datetime.now()
    currentMonth = currentDate.month

    if 9 <= currentMonth <= 12:
        semesterYear = currentDate.year + 1
        semesterAndYear.append("Semester 1, " + str(semesterYear))
        semesterAndYear.append("Semester 2, " + str(semesterYear))
        semesterAndYear.append("Semester 3, " + str(semesterYear))
    
    if 1 <= currentMonth <= 4:
        semestercode = "S2"
        semesterYear = currentDate.year
        semesterAndYear.append("Semester 2, " + str(semesterYear))
        semesterAndYear.append("Semester 3, " + str(semesterYear))
    
    if 5 <= currentMonth <=8:
        semestercode = "S3"
        semesterYear = currentDate.year
        semesterAndYear.append("Semester 3, " + str(semesterYear))

    for i in range(semesterYear+1, 2031):
        for j in range(1, 4):
            semesterAndYear.append("Semester " + str(j) + ", " + str(i))

    combo = ttk.Combobox(labelframe2, width=20, textvariable=comboSemesterYear, state="readonly")
    for i in semesterAndYear:
        combo['values'] = (*combo['values'], i)
        combo.current(0)
    combo.pack()

def createButtonEvent():
    btnOpenSource.bind("<Motion>", lambda event: callback_motion(event,"btn_open_file21_hover.png",btnOpenSource))
    btnOpenSource.bind("<Leave>", lambda event: callback_leave(event,"btn_open_file21.png",btnOpenSource))
    btnGenerate.bind("<Motion>", lambda event: callback_motion(event,"btn_generate_hover.png", btnGenerate))
    btnGenerate.bind("<Leave>", lambda event: callback_leave(event,"btn_generate.png", btnGenerate))
def removeButtonEvent():
    btnOpenSource.unbind("<Motion>")
    btnOpenSource.unbind("<Leave>")
    btnGenerate.unbind("<Motion>")
    btnGenerate.unbind("<Leave>")

window = tk.Tk()
window.iconbitmap(resource_path("favicon.ico"))
window.title("Course Kit Generator")
window.geometry("500x530")
window.configure(bg="white")
window.resizable(FALSE, FALSE)

RadioVariety_1 = StringVar()
comboYear = StringVar()
comboSemesterYear = StringVar()
filePath = StringVar()
fileStatus = BooleanVar()
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

imgOpenSource = PhotoImage(file=resource_path("btn_open_file21.png"))
btnOpenSource = tk.Button(None, text = "button", image = imgOpenSource, 
                command = fileOpen, borderwidth=0,highlightthickness=0)
btnOpenSource.config(justify=CENTER, pady=20)
btnOpenSource.pack()

LabelsFont = font.Font(family='Time New Roman', size=10, weight='bold')
lblFileName = tk.Label(window, wraplength = 450, font=LabelsFont, fg="grey",
                bg="white", textvariable=filePath, borderwidth=0,compound="center",highlightthickness=0)
lblFileName.config(justify=CENTER, pady=10)
lblFileName.pack()

createOptions2()

imgGenerate = PhotoImage(file=resource_path("btn_generate.png"))
btnGenerate = tk.Button(None, text = "button", image = imgGenerate,
                command = createAllFolder, borderwidth=0,highlightthickness=0, pady=0)
btnGenerate.config(justify=CENTER, pady=50)
btnGenerate.pack()

LabelsFont = font.Font(family='Time New Roman', size=10, weight='bold')
lblProgressName = tk.Label(window, wraplength = 1000, font=LabelsFont, fg="red", 
            bg="white", text="In Progress",borderwidth=0, 
            compound="center",highlightthickness=0)
lblProgressName.config(justify=CENTER, pady=120)

createButtonEvent()

canv=tk.Canvas(window, width=800,height=50, bg="blue", bd=0, highlightthickness=0)
canv.place(x = 0, y = 480)
canv.create_rectangle(0, 0, 1001, 50, fill="blue", outline="")

LabelsFont = font.Font(family='Time New Roman', size=12, weight='bold')
lblDeveloper = tk.Label(window, wraplength = 800, font=LabelsFont, fg="white",
                bg="blue", text="@Developed by X1-S3-2019",
                borderwidth=0,compound="center",highlightthickness=0)
lblDeveloper.place(x=120,y=495)

window.mainloop()
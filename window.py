from tkinter import *
from tkinter import filedialog
from tkinter import font
from tkinter import ttk
from pathlib import Path, PureWindowsPath
import os
import tkinter as tk
import WordHandler
import re


firstSubList = ['Assessments','Class Roll','Course Outline','Course Result Summary','Lecture Material','Other Documents','SpreadSheet']
assesmentsSubList = []
assessmentSecondSubList = ['Drafts', 'Moderation Materials', 'Submissions']
assessmentThirdSubList = ['Moderation forms', 'Three Samples']

def fileOpen() :
    openFileName = filedialog.askopenfile(
            filetypes=(("Word file", "*.doc"),("All Files","*.*")),
            title = "Choose a file."
            )
    print(openFileName.name)
    filePath.set(openFileName.name)
    window.mainloop()

def folderCreate(directory):
    try:
        if not(os.path.exists(directory)):
            os.makedirs(directory)
    except OSError:
            print("Failed to create directory!!!!!")

def createSub(rootDirectory, folderList):
    try:
        for i in folderList:
            folderName = rootDirectory+"\\"+i
            if not(os.path.exists(folderName)):
                os.makedirs(folderName)
    except OSError:
            print("Failed to create sub directory!!!!!")

def createWeek1to12(rootDirectory):
    try:
        for i in range(1, 13):
            folderName = rootDirectory+"\\week"+str(i)
            if not(os.path.exists(folderName)):
                os.makedirs(folderName)
    except OSError:
            print("Failed to create week directory!!!!!")

def createAssementsSecondSub(rootDirectory):
    try:
        for i in assessmentSecondSubList:
            folderName = rootDirectory+"\\"+i
            if not(os.path.exists(folderName)):
                os.makedirs(folderName)
    except OSError:
            print("Failed to create week directory!!!!!")

def createAllFolder():
    purePath = str(PureWindowsPath(filePath.get()))
    word = WordHandler.WordHandler()
    extractList = word.extractData(purePath.replace('\\', '\\\\'))
    for i in range(1, len(extractList)):
        print(extractList[i])
        assesmentsSubList.append(removechars(extractList[i]))
        
    rootFolderName = "{code}-{semester}-{year}"
    rootFolderName = rootFolderName.replace('{code}', removechars(extractList[0]))
    rootFolderName = rootFolderName.replace("{semester}", RadioVariety_1.get())
    rootFolderName = rootFolderName.replace("{year}", comboYear.get())

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

def createOptions():
    frame_0 = tk.Frame(window, background="white")
    frame_0.pack()

    frame_1 = tk.Frame(frame_0)
    frame_1.pack(expand=True, side=LEFT, fill='both', padx=50)

    labelframe = tk.LabelFrame(frame_1, text="Semester")
    labelframe.pack()

    RadioVariety_1.set("미선택")

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

window = tk.Tk()
window.title("CourseKitMaker")
window.geometry("800x600")
window.configure(bg="white")
window.resizable(FALSE, FALSE)

RadioVariety_1 = StringVar()
comboYear = StringVar()
filePath = StringVar()

imgLogo=tk.PhotoImage(file="logo.png")
label=tk.Label(window, image=imgLogo, borderwidth=0, highlightthickness=0)
label.config(justify=CENTER, pady=150)
label.pack()

LabelsFont = font.Font(family='Time New Roman', size=20, weight='bold')
lblProgName = tk.Label(window, wraplength = 1000, font=LabelsFont, fg="grey", bg="white", text="Course Kit Generator",borderwidth=0,compound="center",highlightthickness=0)
lblProgName.config(justify=CENTER, pady=20)
lblProgName.pack()

imgOpenSource = PhotoImage(file = "btn_open_source.png")
btnOpenSource = tk.Button(None, text = "button", image = imgOpenSource, command = fileOpen, borderwidth=0,highlightthickness=0)
btnOpenSource.config(justify=CENTER, pady=20)
btnOpenSource.pack()


LabelsFont = font.Font(family='Time New Roman', size=12, weight='bold')
lblFileName = tk.Label(window, wraplength = 1000, font=LabelsFont, fg="grey", bg="white", textvariable=filePath, borderwidth=0,compound="center",highlightthickness=0)
lblFileName.config(justify=CENTER, pady=20)
lblFileName.pack()

createOptions()

imgGenerate = PhotoImage(file = "btn_generate.png")
btnGenerate = tk.Button(None, text = "button", image = imgGenerate, command = createAllFolder, borderwidth=0,highlightthickness=0)
btnGenerate.config(justify=CENTER, pady=100)
btnGenerate.pack()

canv=tk.Canvas(window, width=800,height=50, bg="blue", bd=0, highlightthickness=0)
canv.place(x = 0, y = 550)
canv.create_rectangle(0, 0, 1001, 50, fill="blue", outline="")

LabelsFont = font.Font(family='Time New Roman', size=12, weight='bold')
lblDeveloper = tk.Label(window, wraplength = 800, font=LabelsFont, fg="white", bg="blue", text="@Developed by Karin Saleh, WooHyeon Seong, Sangik Lee, Heena Sood",borderwidth=0,compound="center",highlightthickness=0)
lblDeveloper.place(x=120,y=565)


window.mainloop()




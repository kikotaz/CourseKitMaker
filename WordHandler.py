import win32com.client as win32
from pathlib import Path, PureWindowsPath
import os
import re
import ctypes

# Class which will handle all loading data from the .doc file
class WordHandler:   

    # initializer
    def __init__(self):
        print('call WordHandler initializer')
    
    def closeWord(self, wordApp):
            wordApp.Documents.Close()
            wordApp.Quit()

    # Method to check if the uploaded file is a Course Descriptor
    def checkCourseDescriptor(self, filePath):      
        fileStatus = False
        try:
            wordApp = win32.gencache.EnsureDispatch('Word.Application')
            wordApp.Visible = False
            wordDoc = wordApp.Documents.Open(filePath) # open file
        except Exception as e:
            print(e)
            errorBox = ctypes.windll.user32.MessageBoxW
            errorBox(None, 'Wrong Course Descriptor file type or no Descriptor loaded. '
            + 'Please load a Course Descriptor in .doc format and try again.',
            'Wrong or missing Descriptor.', 0)

        try:
            table = wordDoc.Tables(1)
            courseCodeCell = table.Cell(Row = 2, Column = 1).Range.Text
            courseCodeContent = table.Cell(Row = 2, Column = 2).Range.Text
            courseTitleCell = table.Cell(Row = 3, Column = 1).Range.Text
            courseTitleContent = table.Cell(Row = 3, Column = 2).Range.Text
            
            pattern = re.compile('[^a-zA-Z0-9]')

            #If the second field is not Course Code or empty field
            if not pattern.sub('', courseCodeCell) == 'CourseCode' or \
                len(pattern.sub('', courseCodeContent)) < 1:
                raise Exception('Course Code')
            #If the third field is not Course Title or empty field
            elif not pattern.sub('', courseTitleCell) == 'CourseTitle' or \
                len(pattern.sub('', courseTitleContent)) < 1:
                raise Exception('Course Title')
            else:
                print('Correct file format')
                fileStatus = True
                self.closeWord(wordApp)
        except Exception as e:
            errorBox = ctypes.windll.user32.MessageBoxW
            errorBox(None, 'The file you have chosen is not in proper Course Descriptor format ' 
                + 'or the ' + str(e) +' field is empty. Please check the file and try again.',
                'Course Descriptor file error', 0)
            self.closeWord(wordApp)
        
        print('File Status == ' + str(fileStatus))
        return fileStatus

    # Method to extract all the required data from the word file
    def extractData(self, filePath):
        
        try:
            wordApp = win32.gencache.EnsureDispatch('Word.Application')
            wordApp.Visible = False
            wordDoc = wordApp.Documents.Open(filePath) # open file
        except Exception as e:
            print(e)
            errorBox = ctypes.windll.user32.MessageBoxW
            errorBox(None, 'Wrong Course Descriptor file type or no Descriptor loaded. '
            + 'Please load a Course Descriptor in .doc format and try again.',
            'Wrong or missing Descriptor.', 0)

        # Extracting the course code from the file
        table = wordDoc.Tables(1)
        courseCode = table.Cell(Row = 2, Column = 2).Range.Text
        courseTitle = table.Cell(Row = 3, Column = 2).Range.Text

        # Method to get the row index of any cell by searching its contents
        def getIndex(cellConent):
            for cell in table.Range.Cells:
                if cellConent in cell.Range.Text:
                    return cell.RowIndex

        # Retrieving the list of assessments from file
        assessmentIndex = getIndex('Summative Assessment')
        contentIndex = getIndex('Content')
        assessments = list()
        assessments.append(courseCode)
        assessments.append(courseTitle)
        for row in range(assessmentIndex + 1, contentIndex):
            cellConent = table.Cell(row, 2).Range.Text
            assessments.append(cellConent)
        print(assessments)
        self.closeWord(wordApp)
        return assessments

    def createOutline(self, courseCode, title, sem, year, filePath):

        def replaceTextInHeader(wordDoc, old, new):
            for section in wordDoc.Sections:
                wordDoc.TrackRevisions = False
                headers = section.Headers
                for header in headers:
                    headerRange = header.Range
                    headerRange.Find.Execute(old, ReplaceWith = new, MatchWholeWord = True)
        try:
            # Loading Word Application and opening specified Word file
            outlineTemplate = os.path.dirname(os.path.abspath(__file__)) + '\\CourseOutline.docx'
            print(outlineTemplate)
            wordApp = win32.gencache.EnsureDispatch('Word.Application')
            wordApp.Visible = False
            wordDoc = wordApp.Documents.Open(outlineTemplate)
            wordApp.Selection.Find.Execute('[COURSE]', ReplaceWith = courseCode + ' ' + title,
                MatchWholeWord = True)
            replaceTextInHeader(wordDoc, '[COURSE]', courseCode + ' ' + title)
            replaceTextInHeader(wordDoc, '[SEM]', sem)
            replaceTextInHeader(wordDoc, '[YEAR]', year)

            outputPath = filePath + '\\' + courseCode + '-' + sem + '-' + year + '-' + 'CourseOutline-draft0.docx'

            print(outputPath)
            wordApp.ActiveDocument.SaveAs(outputPath)
            wordApp.Documents.Close()
            wordApp.Quit()
            return "S"
        except Exception as e:
            print(e)
            wordApp.Documents.Close()
            wordApp.Quit()
            return "F"
    
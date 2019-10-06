import win32com.client as win32
import os
import re
import ctypes

# Class which will handle all loading data from the .doc file
class WordHandler:
    # initializer
    def __init__(self):
        print('call WordHandler initializer')
    
    # Method to extract all the required data from the word file
    def extractData(self, filePath):
        try:
            # Loading Word Application and opening specified Word file
            wordApp = win32.gencache.EnsureDispatch('Word.Application')
            wordApp.Visible = False
            wordDoc = wordApp.Documents.Open(filePath)           
        except Exception:
            errorBox = ctypes.windll.user32.MessageBoxW
            errorBox(None, 'Wrong Course Descriptor file type or no Descriptor loaded. '
            + 'Please load a Course Descriptor in .doc format and try again.',
            'Wrong or missing Descriptor.', 0)

        # Extracting the course code from the file
        table = wordDoc.Tables(1)
        courseCode = table.Cell(Row = 2, Column = 2).Range.Text

        # Method to get the row index of any cell by searching its contents
        def getIndex(cellConent):
            for cell in table.Range.Cells:
                if cellConent in cell.Range.Text:
                    return cell.RowIndex
        
        def closeWord():
            wordApp.Documents.Close()
            wordApp.Quit()
            successBox = ctypes.windll.user32.MessageBoxW
            successBox(None, 'Course Kit created successfully', 'Congratulations', 0)

        # Retrieving the list of assessments from file
        assessmentIndex = getIndex('Summative Assessment')
        contentIndex = getIndex('Content')
        assessments = list()
        assessments.append(courseCode)
        for row in range(assessmentIndex + 1, contentIndex):
            cellConent = table.Cell(row, 2).Range.Text
            assessments.append(cellConent)
        print(assessments)
        closeWord()
        return assessments
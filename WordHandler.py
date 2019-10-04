import win32com.client as win32
import os
import re


class WordHandler:
    # initializer
    def __init__(self):
        print('call WordHandler initializer')
    
    def extractData(self, filePath):
        wordApp = win32.gencache.EnsureDispatch('Word.Application')
        wordApp.Visible = True

        wordDoc = wordApp.Documents.Open(filePath)

        table = wordDoc.Tables(1)

        courseCode = table.Cell(Row = 2, Column = 2).Range.Text

        def getIndex(cellConent):
            for cell in table.Range.Cells:
                if cellConent in cell.Range.Text:
                    return cell.RowIndex

        assessmentIndex = getIndex('Summative Assessment')
        contentIndex = getIndex('Content')

        assessments = list()
        assessments.append(courseCode)
        for row in range(assessmentIndex + 1, contentIndex):
            cellConent = table.Cell(row, 2).Range.Text
            assessments.append(cellConent)

        print(assessments)
        return assessments
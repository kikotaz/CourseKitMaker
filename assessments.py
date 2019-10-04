import win32com.client as win32
import os

wordApp = win32.gencache.EnsureDispatch('Word.Application')
wordApp.Visible = True

wordDoc = wordApp.Documents.Open('C:\\Users\\Admin\\Desktop\\COMP504-sample.doc')

table = wordDoc.Tables(1)

courseCode = table.Cell(Row = 2, Column = 2).Range.Text
print(courseCode)

def getIndex(cellConent):
    for cell in table.Range.Cells:
        if cellConent in cell.Range.Text:
            return cell.RowIndex

assessmentIndex = getIndex('Summative Assessment')
contentIndex = getIndex('Content')

assessments = list()
for row in range(assessmentIndex + 1, contentIndex):
    cellConent = table.Cell(row, 2).Range.Text
    assessments.append(cellConent)

print(assessments)
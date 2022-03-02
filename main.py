import pathlib
import win32com.client as win32

# Default Word replace has a 256 character limit for replace text
# this function splites the replace text to 200 character slices
# and replaces the old text with new one
def replace_string(doc, old, new):
    new = str(new)
    n = 200
    chunks = [new[i:i+n]+old for i in range(0, len(new), n)]
    chunks.append("")
    for chunk in chunks:
        doc.Content.Find.ClearFormatting()
        #doc.Content.Find.Replacement.ClearFormatting()
        doc.Content.Find.Execute(old, True, False, False, False, False, False, 1, False, chunk, 2)

# Dispatch an Excel application
excel = win32.Dispatch("Excel.Application")
# Dispatch a Word application
word = win32.Dispatch("Word.Application")

# Get the current path
path = str(pathlib.Path(__file__).parent.resolve())

# Open the excel file
workbook = excel.Workbooks.Open(path + r'\excels\test.xlsx')
sheet    = workbook.Sheets(1) #get the first sheet

for i,row in enumerate(sheet.UsedRange.Rows):
    if(i==0): continue
    name        = row.Cells[1]
    continent   = row.Cells[2]
    capital     = row.Cells[3]
    language    = row.Cells[4]
    population  = row.Cells[5]
    es_date     = row.Cells[6]
    new_year    = row.Cells[7]
    calendar    = row.Cells[8]
    
    # Check all pictures in the excel
    for img in sheet.Shapes:
        # check if the image is related to this row
        # cell 0 is where there is a picture
        if(img.TopLeftCell.Address == row.Cells[0].Address):
            # prepare the picture for the Word file
            # the width of the picture in the Word file must be 100
            img.Width = 100
            # copy the picture to the clipboard
            img.CopyPicture()
    
    # Open the Word template file
    doc = word.Documents.Open(path+r'\templates\1.docx')
    
    # replace the values with the keywords in the template
    replace_string(doc,"NAME",name)
    replace_string(doc,"CONTINENT",continent)
    replace_string(doc,"CAPITAL",capital)
    replace_string(doc,"LANGUAGE",language)
    replace_string(doc,"POPULATION",population)
    replace_string(doc,"ESTABLISHMENTDATE",es_date)
    replace_string(doc,"NEWYEAR",new_year)
    replace_string(doc,"CALENDAR",calendar)
    replace_string(doc,"FLAG",'^c')
    name = path + '\\out\\' + str(name) + '.docx'
    doc.SaveAs(name)
    doc.Close()
workbook.Close()
excel.Quit()
word.Quit()
    

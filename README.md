# written-word-report-from-excel
This Script Export an Excel File to Written Word Reports

# Requirements
This script uses win32com, so first install win32com for python.
```
pip install pywin32
```

As a result, this script is probably works only on Windows, and you must have installed Microsoft Word and Excel.

# Structure
There are three main directories:  
├── excels  
├── templates  
├── out  
└── main.py   

The scripts produces written Word reports based on the templates in `templates` directory.
You may find a default template `1.docx` in `templates` directory which is a sample text.
A sperate Word file will be created for each row of the excel file.

# Usage
Prepare your excel file. Prepare your templates Word files. run `python main.py`.

# Example
We have a `test.xlsx` file which contains information about countries. The first column of the `test.xlsx` contains the flag of the country.
The other columns other contain information, we get and name each row by this way:
```
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
```

In our template file, we have prepared asample text with indicators like NAME, CONTINENT, ... and we will replace this indicators with the above mentioned values like this:
```
    replace_string(doc,"NAME",name)
    replace_string(doc,"CONTINENT",continent)
    replace_string(doc,"CAPITAL",capital)
    replace_string(doc,"LANGUAGE",language)
    replace_string(doc,"POPULATION",population)
    replace_string(doc,"ESTABLISHMENTDATE",es_date)
    replace_string(doc,"NEWYEAR",new_year)
    replace_string(doc,"CALENDAR",calendar)
```

So you need to prepare you Word templates with your desired indicators. 
It is also a good to prepare multiple templates and use them randomly or in order, so if you have 3 templates you may use something like this:
```
    if(i%2==0):
        doc = word.Documents.Open(path + r'\templates\2.docx')
    elif(i%3==0):
        doc = word.Documents.Open(path + r'\templates\3.docx')
    else:
        doc = word.Documents.Open(path + r'\templates\1.docx')
```

# Test example
We have an Excel file like this:
![excel](https://github.com/shahind/written-word-report-from-excel/blob/5cb9e47b6f427f4f57daf2394eb0c1ceb13b331f/excel.png)

And a template like this:
![template](https://github.com/shahind/written-word-report-from-excel/blob/5cb9e47b6f427f4f57daf2394eb0c1ceb13b331f/word.png)

The result would be like this:
![result](https://github.com/shahind/written-word-report-from-excel/blob/5cb9e47b6f427f4f57daf2394eb0c1ceb13b331f/out.png)

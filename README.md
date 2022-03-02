# تبدیل فایل اکسل به گزارش های متنی ورد
این اسکریپت پایتون یک فایل اکسل را به گزارش های متنی در قالب ورد تبدیل می کند
# پیش نیازها
این اسکریپت از کتابخانه win32com استفاده می کند، بنابراین ابتدا آن را نصب کنید:
```
pip install pywin32
```

به دلیل استفاده از این کتابخانه، احتمالا تنها در سیستم عامل ویندوز قادر به استفاده از این اسکریپت هسنید، در ضمن باید ورد و اکسل را هم روی سیستم خود نصب داشته باشید

# ساختار کد
سه پوشه اصلی در این کد وجود دارد:  
```
├── excels  
├── templates  
├── out  
└── main.py   
```

این اسکریپت بر اساس قالب های موجود در پوشه `templates` متن ها را آماده می کند.
قالب `1.docx` در پوشه `templates` به عنوان قالب نمونه ای برای فایل `test.xlsx` تهیه شده است.
بر اساس این قالب فایل ورد جداگانه ای برای هر سطر از فایل اکسل تهیه خواهد شد.

# نحوه استفاده
فایل اکسل و قالب های خود را آماده کنید و کد `python main.py` را اجرا کنید.

# توضیح کد
فایل `test.xlsx` را در نظر بگیرید. این فایل اطلاعاتی از کشور های محتلف را در خود دارد.
اولین ستون فایل `test.xlsx` تصویری از پرچم کشورها را در خود دارد. خواندن عکس ها در فایل اکسل به سادگی خواندن داده های سطر و ستون ها نیست و این پرچم برای نشان داده نحوه کار با عکس ها در فایل اکسل آورده شده است.
سایر ستون ها نیز اطلاعات دیگری در بر دارند، ابتدا لازم است اطلاعات هر یک از ستون ها را برای یک سطر خوانده و در متغیرهایی ذخیره کنیم:
```python
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

توجه کنید که `Cells[0]` شامل عکس پرچم کشورهاست و همانطور که گفتیم خواندن مقدار آن مشابه داده های معمولی نیست و در اینجا مقدار آن را استفاده نمی کنیم. سطر 0ام نیز هدر یا عنوان ستون هاست و با `continue` از آن عبور کرده ایم زیرا نیازی به مقادیر آن نداریم.

در فایل قالب، ما محل قرار گیری داده ها را با کلماتی با حروف بزرگ انگلیسی مشخص کرده ایم، مانند NAME, CONTINENT, ... بدین ترتیب، به سادگی می توانیم آنها را با مقادیر موجود در هر سطر که در متغیرهای بالا ذخیره کردیم جایگزین کنیم:
```python
    replace_string(doc,"NAME",name)
    replace_string(doc,"CONTINENT",continent)
    replace_string(doc,"CAPITAL",capital)
    replace_string(doc,"LANGUAGE",language)
    replace_string(doc,"POPULATION",population)
    replace_string(doc,"ESTABLISHMENTDATE",es_date)
    replace_string(doc,"NEWYEAR",new_year)
    replace_string(doc,"CALENDAR",calendar)
```

در نتیجه لازم است بنابر فایل اکسل خودتان این متغیر ها و جای آنها را در قالب خودتان مشخص و تعریف کنید.
به منظور ایجاد تنوع نیز می توانید قالب های مختلفی ایجاد کنید و به صورت تصادفی یا با ترتیب مشخص از آنها استفاده کنید تا متن های تولید شده کمتر شبیه یکدیگر باشند، به عنوان مثال اگر 3 فایل قالب تعریف کرده باشید می توانید به صورت زیر از آنها استفاده کنید:
```python
    if(i%2==0):
        doc = word.Documents.Open(path + r'\templates\2.docx')
    elif(i%3==0):
        doc = word.Documents.Open(path + r'\templates\3.docx')
    else:
        doc = word.Documents.Open(path + r'\templates\1.docx')
```

# خروجی های فایل مثال
فایل اکسل نمونه ما محتوایی مانند زیر دارد:
![excel](https://github.com/shahind/written-word-report-from-excel/blob/5cb9e47b6f427f4f57daf2394eb0c1ceb13b331f/excel.png)

قالبی که برای این فایل تعریف کرده ایم مانند زیر است(به کلمات با حروف بزرگ دقت کنید):
![template](https://github.com/shahind/written-word-report-from-excel/blob/5cb9e47b6f427f4f57daf2394eb0c1ceb13b331f/word.png)

پس از اجرا، دو فایل ورد(برابر تعداد سطرهای فایل اکسل) به صورت زیر در پوشه `out` خواهیم داشت:
![result](https://github.com/shahind/written-word-report-from-excel/blob/5cb9e47b6f427f4f57daf2394eb0c1ceb13b331f/out.png)


# Turn Excel files to written Word reports
This python script exports an Excel file to written Word reports.

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

The script produces written Word reports based on the templates in `templates` directory.
You may find a default template `1.docx` in `templates` directory which is a sample template for `test.xlsx`.
A sperate Word file will be created for each row of the excel file.

# Usage
Prepare your excel file. Prepare your Word template files. run `python main.py`.

# Example Description
We have a `test.xlsx` file which contains information about countries. The first column of the `test.xlsx` contains the flag of the country.
The other columns contain other information, we get and name each row by this way:
```python
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
The first row (`i=0`) is the header row which we don't need it. And the `Cell[0]` is the cell which contains the flag picture. Finding the pictures in Excel is not similar to reading the values of the cells, so we do not get the value of the `Cell[0]` here.  
In our template file, we have prepared a sample text with indicators like NAME, CONTINENT, ... and we will replace these indicators with the above mentioned values like this:
```python
    replace_string(doc,"NAME",name)
    replace_string(doc,"CONTINENT",continent)
    replace_string(doc,"CAPITAL",capital)
    replace_string(doc,"LANGUAGE",language)
    replace_string(doc,"POPULATION",population)
    replace_string(doc,"ESTABLISHMENTDATE",es_date)
    replace_string(doc,"NEWYEAR",new_year)
    replace_string(doc,"CALENDAR",calendar)
```

So you need to prepare your Word templates with your desired indicators for your Excel file. 
It is also a good idea to prepare multiple templates and use them randomly or in order, so if you have 3 templates you may use something like this:
```python
    if(i%2==0):
        doc = word.Documents.Open(path + r'\templates\2.docx')
    elif(i%3==0):
        doc = word.Documents.Open(path + r'\templates\3.docx')
    else:
        doc = word.Documents.Open(path + r'\templates\1.docx')
```

# Test Example
We have an Excel file like this:
![excel](https://github.com/shahind/written-word-report-from-excel/blob/5cb9e47b6f427f4f57daf2394eb0c1ceb13b331f/excel.png)

And a template like this:
![template](https://github.com/shahind/written-word-report-from-excel/blob/5cb9e47b6f427f4f57daf2394eb0c1ceb13b331f/word.png)

The result would be two Word files in `out` directory like this:
![result](https://github.com/shahind/written-word-report-from-excel/blob/5cb9e47b6f427f4f57daf2394eb0c1ceb13b331f/out.png)

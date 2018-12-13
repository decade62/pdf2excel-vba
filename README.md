# pdf2excel-vba
# IGNORE AS THIS IS A SIDE PROJECT.
# Branch version for generating .xls files from .pdf specific reports
User Interface in MS Excel that automates the transfer of data from pdf to the spreadsheet.
The concept is that the text from a pdf file is copied and splitted to rows (by newlines) and columns (by delimiter).

Points to consider:

- The program detects and defaults to the first available pdf file that exists in the current directory. The Folder and File name may change through cell input.
- The buttons provide interface assistance. Even when "Hide Interface" is clicked, it will show again next time the file is opened.
- User may manually change the delimiter that splits text into cells by pressing the corresponding button.

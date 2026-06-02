1. Rename config_template.py to config.py
2. Fill out config.py
3. Move whole "Templates" folder to a directory one above the python 
(E.g.   Invoicing/
        ----> Templates/
        ----------> filled out template files
        ----> Invoice Code/
        ----------> GenerateInvoiceAndTimesheet.py
)
4. Fill out templates in template folder
5. Rename RunLastMonth copy.bat -> RunLastMonth.bat
6. Configure RunLastMonth.bat to the correct windows path of GenerateInvoiceAndTimesheet.py
7. Every month double click RunLastMonth.bat
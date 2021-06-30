# Automation Project with Python
=> Automating tasks in EXCEL Spreadsheet!

- There is an error(human) in the price & we have to decrease each price by 10% 
(i.e we can multiply each price by 0.9)

- The program will automate the process & it will also add a bar-chart to our data.
________________________________________

- installed the 'openpyxl' module with pip (its a Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files.)

- we can set an alias (shortName) for a module to be used in our file, instead of the big name of the module
<import module as aliasName>
=> this will allow to access the module functions through the alias name!

- make sure that while uploading the workbook, you are in the current directory where the workbook is present!
- we should name the workbook object as <wb> (just for understanding did otherwise)

- to know the number of rows in sheet:
<sheetObj.max_row>

- to know the number of columns in sheet:
<sheetObj.max_column> 

- if we don't add 1 in the range of the loop, then rows will be generated from 1 to (max_row-1)!
- sheet.cell(row,3) -> means that the concerned row & in that the 3rd column (which has all the price values)

- <fileObj.save('newFileName')> -> saves the file that you accessed from the fileObj in a new file! (did this to not over-write the original file, incase of a bug)

=> MAIN TASK IS DONE!
____________________________________________

# Adding a Chart:

- we referenced the values we wanted to use
- we added these values to the BarChart instance to create a chart out of it
- then, we added this chart to our excel file specifying the cellName

# Cleaning uo the code a little:

- we can define a function which takes the file as an input!
- 




















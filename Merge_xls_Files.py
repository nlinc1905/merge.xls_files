"""

Merge all xls files within a directory into 1 xls file with multiple sheets.
Each sheet is named according to the filename of the xls file merged.
Sheet names are input as substrings to stay within character limit.

"""

import xlrd, xlwt
import os

#Choose the directory for the input and output files
directory = input("Input the directory of the .xls files to merge, making sure to match spelling, spaces, and letter case exactly: ")
j = os.path.isdir(directory)
while j == False:
    print(directory + " is an invalid directory.")
    directory = input("Please enter a valid directory: ")
    j = os.path.isdir(directory)

#Choose the output file name
j = 0
while j == 0:
    outputFileName = input("Enter your desired output file name: ")
    if len(outputFileName) > 0:
        j += 1

#Build list of .xls files in the directory and create the sheet names
files2merge = []
sheetnames = []
for file in os.listdir(directory):
    if file.endswith(".xls"):
        files2merge.append(file)
        sheetnames.append(file[:-4])
if len(files2merge) == 0:
    print("No .xls files found in the directory specified.")
else: 
    print(files2merge)

#Open a blank .xls file in write mode and store it as a variable
merged_file = xlwt.Workbook()

#Build the merged file
for i,file in enumerate(files2merge):
    #Add a new sheet named after the file being merged
    worksheet = merged_file.add_sheet(sheetnames[i][:30])
    #Read the file being merged
    openBook = xlrd.open_workbook(directory + "/" + file)
    #Write the data from the 1st sheet of the file being merged to the new file
    for sheet in range(openBook.nsheets):
        openSheet = openBook.sheet_by_index(sheet)
        for rx in range(openSheet.nrows):
            for cx in range(openSheet.ncols):
                worksheet.write(rx, cx, openSheet.cell_value(rx, cx))

#Save the merged .xls file to the directory
merged_file.save(directory + "/" + outputFileName + ".xls")

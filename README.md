# Python---How-to-Merge-all-excel-files-in-a-folder-to-single-excel-file
Python - How to Merge all excel files in a folder

To merge all excel files in a folder, use the Glob module and the append() method.

Let’s say the following are our excel files on the Desktop −

Sales1.xlsx



Sales2.xlsx



Note − You may need to install openpyxl and xlrd packages.

At first, set the path where all the excel files you want to merge are located. Get the excel files and read them using glob −

path = "C:\\Users\\amit_\\Desktop\\"

filenames = glob.glob(path + "\*.xlsx")
print('File names:', filenames)
Next, create an empty dataframe for the merged output excel file that will get the data from the above two excel files −

outputxlsx = pd.DataFrame()
Now, the actual process can be seen i.e. at first, iterate the excel files using for loop. Read the excel files, concat them and append the data −

for file in filenames:
   df = pd.concat(pd.read_excel(file, sheet_name=None), ignore_index=True, sort=False)
   outputxlsx = outputxlsx.append(df, ignore_index=True)
Example
Following is the code −

import pandas as pd
import glob

# getting excel files to be merged from the Desktop 
path = "C:\\Users\\amit_\\Desktop\\"

# read all the files with extension .xlsx i.e. excel 
filenames = glob.glob(path + "\*.xlsx")
print('File names:', filenames)

# empty data frame for the new output excel file with the merged excel files
outputxlsx = pd.DataFrame()

# for loop to iterate all excel files
for file in filenames:
   # using concat for excel files
   # after reading them with read_excel()
   df = pd.concat(pd.read_excel( file, sheet_name=None), ignore_index=True, sort=False)

   # appending data of excel files
   outputxlsx = outputxlsx.append( df, ignore_index=True)

print('Final Excel sheet now generated at the same location:')
outputxlsx.to_excel("C:/Users/amit_/Desktop/Output.xlsx", index=False)
Output
This will produce the following output i.e., the merged excel file will get generated at the same location −


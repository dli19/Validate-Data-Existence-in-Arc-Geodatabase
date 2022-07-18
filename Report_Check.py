## Report Sample: June 2021 ##


import arcpy, os, xlrd

#Assign existing GIS feature file path to variable "adp"
adp = "C:/GIS/Master_Working/Master_Working.gdb/ADCOM_SSAP_MERGED_SCHEMA"

#Make sure the Excel file is saved in xls format (As in July, 2022)
report = 'June 2021.xls'

#Assign newest monthly report excel file path to variable "fileL"
fileL = f"F:/Jim Li (jli)/Adam Monthly Planning Report/{report}"

#Apply open_workbook funtion from the xlrd library, and assign the Excel file to a variable named "wb"
wb = xlrd.open_workbook(fileL)

#Apply sheet_by_index function from the xlrd library, and assign the value to a variable named "sheet"
#Sheets in Microsoft Excel are starting with index 0
#The sheet to be processed is in sheet1 of the excel file, in this case its index is 1
sheet = wb.sheet_by_index(1)

#Create three empty lists for later data stored
coLis = []
preLis = []
feaLis = []

#To index the attributes in Excel by rows (vertically), 
#Starting with index 1, because first row is field name
#Cell_value function finds value in i row of the 4th field from the left, in this case E field of the Excel sheet
#For each value that foud in each row of 4th field, append the value to empty list coLis[]
#Apply upper() function to standardized the value to ensure consistency 
for i in range(1,sheet.nrows):
    coLis.append(sheet.cell_value(i,4).upper())


#Use an Arcpy search cursor to find value in the feature class's attribute table
#Look for addresses that are in the FullAddress field of the feature class
#Then append each value to the empty list preLis[], (watch out for the data type!! in this step)
with arcpy.da.SearchCursor(adp, ['FullAddress']) as cursor:
    for row in cursor:
        preLis.append(row)


#Convert feature class attribute values to a python list data type, therefore append any value in preLis[] to feaLis[]
for a in preLis:
    feaLis.append(a[0])

    

#Verify the number of features in Excel file and the number of address in the GIS dataset
print(f'The total number of address points in ADCOM_SSAP is',len(feaLis))
print(f'The total number of addresses for the respective month of the Excel sheet is',len(coLis))


#Validate Data Existence between the Excel file and GIS feature dataset
check = all(item in coLis for item in feaLis)

#Conditional check
if check is True:
    print('Nothing new need to be updated!')
else:
    print('There are new addresses to be added, and they are listing below!!')


#print out any addresses that are not in the GIS address point dataset but in the Excel file
for item in coLis:
    if item not in feaLis:
        print(item)

#Conclude the operation is completed
print('Done!')
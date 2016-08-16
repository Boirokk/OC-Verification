# Written by Boirokk 2016-08-16
#1.2 added blanks h_row.append('') to data and headers

import xlrd
import os
import csv


# Look for query in each sheet and print rows containing query
def find_in_workbook_sheets(file_location,new_file_name):

    # Open file and store in workbook variable
    workbook = xlrd.open_workbook(file_location)
   

    # Open each sheet, search for criteria entered by user and write matching rows to new file
    sheet = workbook.sheet_by_index(0) #Open worksheet in workbook
    h_row = [] #Create list

    # Add cell value to list in order
    h_row.append(sheet.cell_value(7, 1))
    h_row.append('')
    h_row.append('')
    h_row.append(sheet.cell_value(5, 1))
    h_row.append('')
    h_row.append('')
    h_row.append(sheet.cell_value(0, 1))
    h_row.append(sheet.cell_value(51, 1))
    h_row.append('')
    h_row.append(sheet.cell_value(34, 1))
    h_row.append(sheet.cell_value(25, 1))
    h_row.append(sheet.cell_value(20, 1))

    h_row.append('')
    h_row.append(sheet.cell_value(35, 1))
    h_row.append('')
    h_row.append('')
    h_row.append(sheet.cell_value(26, 1))
    h_row.append(sheet.cell_value(3, 1))
    h_row.append('')
    h_row.append('')
    h_row.append('')
    h_row.append('')
    
    # Write to csv file in order
    with open(new_file_name,'a',newline='') as fp:
        a =  csv.writer(fp,delimiter=',')
        data = [h_row]
        a.writerows(data)
        


# Get the .xls and .xlsx files from the root and sub dirs
def get_file_contents(new_file_name):

    
    file_location = r'C:\Users\comms\Desktop\WO CHECK'

    for roots, dirs, files in os.walk(file_location):
        for file in files:
            file_name = roots + '\\' + file
           
            if '.xlsx' in file_name:
                try:
                    find_in_workbook_sheets(file_name,new_file_name)
                except:
                    continue
            elif '.xls' in file_name:
                try:
                    find_in_workbook_sheets(file_name,new_file_name)
                except:
                    continue
            
# main
def main():

    new_file_name = "OC Verification.csv"
    print('Creating file... ', new_file_name)
    
    # Open new .xls file and insert headers
    try:
        with open(new_file_name,'w',newline='') as fp:
            a =  csv.writer(fp,delimiter=',')
            data = [['CUSTOMER','DPRO SO#','PART #','WO#', '', 'Rev','DATE', 'FT','MATERIAL', 'MANUF AND MODEL', 'NAME/HULL#', 'PO#','QUOTE', 'AREA','LIN. FT.', 'T/M', 'BOATYARD', 'REP.','% oF COMP', 'TOTAL FT', 'STATUS', 'EST TIME']] # Create Headers
            a.writerows(data)
    except:
        print('Please close the Accumulated Data Search Document and try again.')
        error = input('Press enter to exit')
        exit()


    get_file_contents(new_file_name)
    print('Done...')
    input('Press enter to exit')
    
# Call main
main()

from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
wb = load_workbook(filename='Holiday_Inn_Al_Barsha_2020_old.xlsx')
# output sheet
output_sheet = wb['MMTB']
max_output_row = output_sheet.max_row
# getting column number from name : https://stackoverflow.com/a/12902801
xy = coordinate_from_string('AM4') # returns ('A',4)
col = column_index_from_string(xy[0])
xy = coordinate_from_string('AF4') 
out_col = column_index_from_string(xy[0]) 
print(out_col)
input_sheet = wb['CTB 20']
max_input_row = input_sheet.max_row
#output_sheet['AF4'] = 'BLAH BLAH'
output_sheet.cell(row=4, column = 32).value = "=\'CTB 20\'!AM13"

  
for i in range(1,max_output_row+1):
    cell_out = output_sheet.cell(row = i, column = 2)
    for j in range(1,max_input_row+1):
        cell = input_sheet.cell(row = j, column = 2)
        if cell.value != None:
            if cell_out.value == cell.value:
                output_sheet.cell(row=i,column=out_col).value = '=\'CTB 20\'!AM{}'.format(j)
    #wb.save("outputfile.xlsx")  
wb.save("outputfile.xlsx")   
           

        

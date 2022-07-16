from unittest import skip
from openpyxl import load_workbook, Workbook     
from openpyxl.utils import get_column_letter
 

wb= load_workbook ('Pack.xlsx') # Or name of your file.xlsx
ws = wb.active

partials = 0
full = 0
#Pallet Count
for row in range (5, 900):
    for col in range(10, 11):
        char = get_column_letter(col) 
        selected_box = ws[char + str(row)].value  
        
        if selected_box == None:
            skip            
        else:
            pallets = str(selected_box)
            if selected_box == "PARTIAL":
                partials = partials +1
            if selected_box == "FULL":
                full = full +1
            

ranch = 0
last = 0
#Last Pallet Received 
for row in range(5, 900):
    for col in range(7, 8):
        char = get_column_letter(col)
        selected_box = ws[char + str(row)].value 
        
        if selected_box == None :
            skip
        else:
            time = str(selected_box).split(" ")[-1]
            time = int(time.replace(":", ""))
        if time > last:
            last = time
        
        
for row in range(5, 900):
    for col in range(1, 2):
        char = get_column_letter(col)
        selected_box = ws[char + str(row)].value
        if selected_box == None :
            skip
        else:
            ranch = selected_box
            
                    
print(f"""\n---------------------------------------------
There are {full} Full and {partials} Partial Pallets 
Last pallet received at {last} by Ranch {ranch} 
---------------------------------------------- \n""")

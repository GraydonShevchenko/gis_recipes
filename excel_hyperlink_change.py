#description: 
# this script updates a path in an excel document 
import openpyxl 

import os
#adding workbooko
wb = openpyxl.load_workbook(filename=r"")
#pointing to the worksheet script is working on
sheet = wb.worksheets[2]

#dictionary to create the switch 
#format is - old_root : new root 
new_roots = { r"":
             r""}

# iterate through all cells containing hyperlinks
for row in sheet.iter_rows():
    for cell in row:
        if cell.hyperlink:
            target = cell.hyperlink.target
            #only want the diractory path 
            target_head = os.path.dirname(target)
            #iterate through the new roots items
            for old_root, new_root in new_roots.items(): 
                #match target directory and old root directory path 
                if target_head == old_root: 
                    #replace it with the new root
                    new_target = target.replace(old_root,new_root)
                    #update the spreadsheet 
                    cell.hyperlink.target = new_target
                    #print it out
                    print(new_target)
                else: 
                    #print those who don't make it 
                    print(old_root)
#save a new workbook for version control 
wb.save(r"")



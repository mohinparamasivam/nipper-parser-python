from bs4 import BeautifulSoup as bs
import time
import xlwt 
from xlwt import Workbook 

content = []
# Read the XML file
with open("seremban.xml", "r") as file:
    # Read each line in the file, readlines() returns a list of lines
    content = file.readlines()
    # Combine the lines in the list into a string
    content = "".join(content)
    bs_content = bs(content, "lxml")


find_anchor_counter = 0

found_anchor = False

print("\n")
sect_start_input = int(input("Enter start <sect1 location number : ")) # find the start <sect1
sect_end_input = int(input("Enter <sect1 location number : ")) # find the end <sect1


sect_start = sect_start_input-1
sect_end = sect_end_input-1
	
section_part_column = 0
section_part_row = 0

wb = Workbook() 
sheet1 = wb.add_sheet('Sheet 1')


# While there is more data in the xml file continue looping
while(sect_start!=sect_end):
	
	
	#If the column reaches 3 then reset column and row coordinates
	
	if(section_part_column==4):
		section_part_column=0
		section_part_row+=1
		content = bs_content.findAll('sect1')[sect_start].text
		print("Writing location : Row --> %s Column--> %s " %(section_part_row,section_part_column))
		sheet1.write(section_part_row,section_part_column,content)
		wb.save('test.xls')
		section_part_column+=1
		
		
	else:
	
		content = bs_content.findAll('sect1')[sect_start].text
		print("Writing location : Row --> %s Column--> %s " %(section_part_row,section_part_column))
		sheet1.write(section_part_row, section_part_column,content)
		wb.save('test.xls')
		section_part_column+=1
		
		
		
	#Counter for sect incement
	sect_start+=1
	
			
			
			
			
	
		
		
		
		
		
		
		
		
		
		
		
		

from bs4 import BeautifulSoup as bs
import xlwt 
from xlwt import Workbook 
import xlsxwriter





print ("                                                                       ")
print ("                                                                       ")
print ("            #####################################################" )       
print ("            #                                                   #" )        
print ("            #              Nipper XML Report Parser             #" ) 
print ("            #                                                   #" )                            
print ("            #            Author : Mohin Paramasivam             #" )       
print ("            #                                                   #" )       
print ("            #####################################################" )    


content = []
# Read the XML file

print("\r\n")
print("   Convert HTML Report file to XML at https://www.freefileconvert.com\r\n\r\n")
read_file_xml = input("Nipper XML Input File : ")
print("\r\n")
output_file_excel = input("Enter Output Filename : ")

with open(read_file_xml, "r") as file:
    # Read each line in the file, readlines() returns a list of lines
    content = file.readlines()
    # Combine the lines in the list into a string
    content = "".join(content)
    bs_content = bs(content, "lxml")


find_anchor_counter = 0

found_anchor = False

print("\r\n")
sect_start_input = int(input("Enter start <sect1> location number : ")) # find the start <sect1>
print("\r\n")
sect_end_input = int(input("Enter end <sect1> location number : ")) # find the end <sect1>


sect_start = sect_start_input-1
sect_end = sect_end_input-1
	
section_part_column = 0
section_part_row = 0

wb = xlsxwriter.Workbook(output_file_excel+".xls")
sheet1 = wb.add_worksheet()


# While there is more data in the xml file continue looping
while(sect_start!=sect_end):
			
	
	#If the column reaches 3 then reset column and row coordinates
	
	if(section_part_column==4):
		section_part_column=0
		section_part_row+=1
		content = bs_content.findAll('sect1')[sect_start].text
		print("[+] Writing location : Row --> %s Column--> %s " %(section_part_row,section_part_column))
		sheet1.write_string(section_part_row,section_part_column,content)
		section_part_column+=1
		
		
	else:
	
		content = bs_content.findAll('sect1')[sect_start].text
		print("[+] Writing location : Row --> %s Column--> %s " %(section_part_row,section_part_column))
		sheet1.write_string(section_part_row, section_part_column,content)
		section_part_column+=1
		
		
		
	#Counter for sect increment
	sect_start+=1
	
print("\r\n")
print("[+] Successfully Written to --> %s" %(output_file_excel))	
print("\r\n")	
		
		
		
		
		
		
		
		
		
		
		

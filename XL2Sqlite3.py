import sqlite3
import xlrd
import os

path = '/home/greg/Documents/'
path_xls = path + 'xls_files/'
rugby = [ ]
club = "RUFC_Sydenham.xlsx"

'''wb = xlrd.open_workbook(path_xls + club)
worksheets = wb.sheet_names()
print (wb.sheet_names())
conn = sqlite3.connect(path + 'Sqlite3 Databases/RUFC_Sydenham.db')
'''


for filename in os.listdir(path_xls):
	if "RUFC" in filename:
		rugby.append(filename)
		print rugby
		
		
for club in rugby:
	wb = xlrd.open_workbook(path_xls + club)
	#remove everything after the '.' for the sqlite3 filename!!!
	worksheets = wb.sheet_names()
	print ( wb.sheet_names() )
	pos = club.find('.')
	print(pos)
	club = club[:pos]
	conn = sqlite3.connect(path + 'Sqlite3 Databases/' + club + '.db')
	c = conn.cursor()
		

	for sh_name in worksheets:
		print (sh_name)
		curr_row = 0
		worksheet = wb.sheet_by_name(sh_name)
		num_cols = worksheet.ncols
		num_rows = worksheet.nrows
		qmarks = "("

		for i in range(num_cols - 1):
			qmarks += "?, "

		qmarks += "?)"
		vals = sh_name + " Values " + qmarks
		print("WS name & values:  " + vals)

		while curr_row < num_rows:
			row = worksheet.row(curr_row)
			curr_col = 0
			hdr = " "
			line = [ ]
			xline = [ ]
			
			while curr_col < num_cols:
			#Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
			#cell_type = worksheet.cell_type(curr_row, curr_col)
				cell_type = worksheet.cell_type(curr_row, curr_col)
				print cell_type
				cell_value = worksheet.cell_value(curr_row, curr_col)
				
				if (cell_type == 2):
					cell_value = int(cell_value)
				cell_value = str(cell_value)

				if curr_row == 0:   #This is the Hdr-row - need brackets '(...)'
					if curr_col < num_cols - 1:
						if curr_col == 0:       hdr += "("
						hdr += cell_value + " text, "
					else:   hdr += cell_value + " text)"

				else:   #current row != 0.
					
					if curr_col < num_cols - 1:
						hdr += cell_value + ","
						line.append(cell_value)
					else:
						hdr += cell_value
						line.append(cell_value) 

				if curr_col == num_cols - 1:
					print ("line = " + hdr)
					tuline = tuple(line)
					xline.append(tuline)

					if curr_row == 0:
						c.execute("CREATE TABLE IF NOT EXISTS " + sh_name + hdr)
					else:
						c.executemany('insert into ' + vals, xline)

				curr_col += 1	    #end inner while
			curr_row += 1	#end outer while

	conn.commit()
	conn.close()


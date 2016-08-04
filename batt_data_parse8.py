
import xlrd
from xlwt import Workbook


def get_data(filelist, single_cell, fileout, shownames, showtemp, cellspace1, cellspace2):
	# DEFINES, these come from the standard Arbin output
	INDEX_COL = 0
	CUR_COL = 6
	VOL_COL = 7
	CAP_MAH_COL = 9
	CAP_PWR_COL = 11
	TEMP_COL = 21
	# Include negative sign here for discharging
	DISCHARGE_CUR = -4

	# Create new workbook to store output in
	wb = Workbook()
	out_file = wb.add_sheet('Sheet1')

	# Initialize row and col within the new excel file
	print_row = 0 
	print_col = 0

	# Testing one or two cells?
	if single_cell:
		v_target = [4.1, 4.0, 3.9, 3.8, 3.7, 0]
		CUTOFF = 2.5
	else:
		v_target = [8.2, 8.0, 7.8, 7.6, 7.5, 0]
		CUTOFF = 5

	capacity_out_col = (len(v_target) + 1)*2 + len(v_target)*cellspace1 + cellspace2
	print (capacity_out_col)
	# # Automatically subtracts .001 for rounding purposes 
	for num in range(len(v_target)):
		v_target[num] = v_target[num] - 0.001

	for num in range(len(filelist)):
		new_dataset = True
		i = 2
		j = 0

		# Every new file, start at left and print file name
		print_col = 0
		filename = filelist[num]
		if shownames:
			out_file.write(print_row, print_col, filename)
			print_col += 1

		# Magic number 1 is from arbin output, should be modified to parse sheet names
		# This only works with .xls files, not .xlsx, need to fix!!!!!!!!!!!
		data_in = xlrd.open_workbook(filename).sheet_by_index(1)

		# Initial values do not matter as long as not null
		cur_cell = data_in.cell(rowx = 0, colx = 0)
		last_cell = data_in.cell(rowx = 0, colx = 0)
		last_jump = 1
		cur_jump = 1

		while True:
			# Iterate through the cells, saving current and previous 
			last_cell = cur_cell
			# If you run out of data, break
			try:
				cur_cell = data_in.cell(rowx = i, colx = CUR_COL)
			except:
				break

			# Look at current, every falling edge, record position, check values
			if last_cell.value == 0 and cur_cell.value < DISCHARGE_CUR + 1:
				last_jump = cur_jump
				cur_jump = i
				if new_dataset == True:

					if showtemp:
						out_file.write(print_row, 1, 'temp = ' + str(round(data_in.cell_value(rowx = cur_jump - 1, colx = TEMP_COL))))
					print_col = 3

					out_file.write(print_row, print_col, data_in.cell_value(rowx = cur_jump - 1, colx = VOL_COL))
					print_col += 1

					out_file.write(print_row, print_col, data_in.cell_value(rowx = cur_jump, colx = VOL_COL))
					print_col += cellspace1

					new_dataset = False

				elif (data_in.cell_value(rowx = cur_jump - 1,  colx = VOL_COL) < 
				v_target[j]	< data_in.cell_value(rowx = last_jump - 1,  colx = VOL_COL)):
					print_col += 1

					out_file.write(print_row, print_col, data_in.cell_value(rowx = last_jump - 1, colx = VOL_COL))
					print_col += 1

					out_file.write(print_row, print_col, data_in.cell_value(rowx = last_jump, colx = VOL_COL))
					print_col += cellspace1

					j += 1

			# Check for new tests
			if (data_in.cell_value(rowx = i, colx = VOL_COL) < CUTOFF and data_in.cell_value(rowx = i, colx = CAP_MAH_COL) != 0 
			and data_in.cell_value(rowx = i, colx = CUR_COL) <= DISCHARGE_CUR + 1):
				# Magic number 29 fits the layout currently being used to display data
				print_col = capacity_out_col

				# Capacity in mAh instead of Ah thus *100
				out_file.write(print_row, print_col, data_in.cell_value(rowx = i, colx = CAP_MAH_COL) * 1000)
				print_col += 1			
				out_file.write(print_row, print_col, data_in.cell_value(rowx = i, colx = CAP_PWR_COL))
				print_col += 1

				# Reset trackers and counters, move down a row
				new_dataset = True
				j = 0
				print_col = 0
				print_row += 1
			i += 1
	wb.save(fileout)

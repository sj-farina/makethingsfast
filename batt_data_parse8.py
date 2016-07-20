
import xlrd
from xlwt import Workbook


# filelist = ['LIR18350_DCHG_4000mA_-20C_0C_25C_55C_6-14-2016_1.xls', 'LIR18350_DCHG_4000mA_-20C_0C_25C_55C_6-14-2016_1.xls', 'LIR18350_DCHG_4000mA_-20C_0C_25C_55C_6-14-2016_1.xls']

def get_data(filelist):
	###############
	# Part 0: User Setup
	# Change file and search values
	###############

	# Files must be in same working directory, check sheet name
	# Enter numbers here to change search values, leave zero at the end
	single_cell = True

	if single_cell:
		v_target = [4.1, 4.0, 3.9, 3.8, 3.7, 0]
		CUTOFF = 2.5
	else:
		v_target = [8.2, 8.0, 7.8, 7.6, 7.5, 0]
		CUTOFF = 5

	###############
	# Part 0: Auto Setup
	# Initialize values, load files
	###############


	# DEFINES
	INDEX_COL = 0
	CUR_COL = 6
	VOL_COL = 7
	CAP_MAH_COL = 9
	CAP_PWR_COL = 11
	TEMP_COL = 21

	# Include negative sign here for discharging
	DISCHARGE_CUR = -4

	wb = Workbook()
	out_file = wb.add_sheet('Sheet1')



	print_row = 0 

	for num in range(len(filelist)):
		filename = filelist[num]
		out_file.write(print_row, 0, filename)
		print_row += 1

		# magic number 1 is from ardent output, should be modded to parse sheet names
		# this only works with .xls files, not .xlsx, need to fix
		data_in = xlrd.open_workbook(filename).sheet_by_index(1)

		# # Automatically subtracts .001 for rounding purposes 
		for num in range(len(v_target)):
			v_target[num] = v_target[num] - 0.001

		# Initial values do not matter as long as not null
		cur_cell = data_in.cell(rowx = 0, colx = 0)
		last_cell = data_in.cell(rowx = 0, colx = 0)
		last_jump = 1
		cur_jump = 1

		# # Generic counters for looping
		i = 2
		j = 0
		print_col = 0

		new_dataset = True
		while True:
			# Iterate through the cells, saving current and previous 
			last_cell = cur_cell
			# If you run out of data, break
			try:
				cur_cell = data_in.cell(rowx = i, colx = CUR_COL)
			except:
				break

			# Check for new tests
			if (data_in.cell_value(rowx = i, colx = VOL_COL) < CUTOFF and data_in.cell_value(rowx = i, colx = CAP_MAH_COL) != 0 
			and data_in.cell_value(rowx = i, colx = CUR_COL) <= DISCHARGE_CUR + 1):
				
				print_col += 6

				# capacity in mAh instead of Ah thus *100
				out_file.write(print_row, print_col, data_in.cell_value(rowx = i, colx = CAP_MAH_COL) * 1000)
				print_col += 1			
				out_file.write(print_row, print_col, data_in.cell_value(rowx = i, colx = CAP_PWR_COL))
				print_col += 1


				new_dataset = True
				j = 0
				print_col = 0
				print_row += 1

			# Look at current, every falling edge, record position, check values
			if last_cell.value == 0 and cur_cell.value < DISCHARGE_CUR + 1:
				last_jump = cur_jump
				cur_jump = i
				if new_dataset == True:

					out_file.write(print_row, 0, 'temp = ' + str(round(data_in.cell_value(rowx = cur_jump - 1, colx = TEMP_COL))))
					print_col += 2

					out_file.write(print_row, print_col, data_in.cell_value(rowx = cur_jump - 1, colx = VOL_COL))
					print_col += 1

					out_file.write(print_row, print_col, data_in.cell_value(rowx = cur_jump, colx = VOL_COL))
					print_col += 1

					new_dataset = False

				elif (data_in.cell_value(rowx = cur_jump - 1,  colx = VOL_COL) < 
				v_target[j]	< data_in.cell_value(rowx = last_jump - 1,  colx = VOL_COL)):
					# print('Voltage = ', v_target[j] + .001)
					print_col += 1

					out_file.write(print_row, print_col, data_in.cell_value(rowx = last_jump - 1, colx = VOL_COL))
					print_col += 1

					out_file.write(print_row, print_col, data_in.cell_value(rowx = last_jump, colx = VOL_COL))
					print_col += 1

					j += 1
			i += 1
		print_row += 1 

	wb.save('output.xls')
# get_data(filelist)
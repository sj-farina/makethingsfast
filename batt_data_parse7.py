
import xlrd

# def get_data(filelist):
# 	print("boop", filelist)

#filelist = ['B014-2_7-19-2016_4A_Switch.xls']

def get_data(filelist):
	###############
	# Part 0: User Setup
	# Change file and search values
	###############

	# Files must be in same working directory, check sheet name
	# Enter numbers here to change search values, leave zero at the end
	v_target = [4.1, 4.0, 3.9, 3.8, 3.7, 0]
	CUTOFF = 2.5



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






	for num in range(len(filelist)):
		filename = filelist[num]

		# magic number 2 is from ardent output, should be modded to parse sheet names
		data_in = xlrd.open_workbook(filename).sheet_by_index(1)


		###############
		# Part 1
		# Find capacity
		###############

		# Search for voltage under CUTOFF to indicate cycle end
		cur_cell = data_in.cell(rowx = 0, colx = 0)
		i = 1
		print("mAh capacity, Power capacity")
		while True:
			try:
				cur_cell = data_in.cell(rowx = i, colx = VOL_COL)
			except: 
				break
			if (cur_cell.value < CUTOFF and data_in.cell_value(rowx = i, colx = CAP_MAH_COL) != 0 
			and data_in.cell_value(rowx = i, colx = CUR_COL) <= DISCHARGE_CUR + 1):
				print ("Temperature = ", round(data_in.cell_value(rowx = i, colx = TEMP_COL)))
				# Toggle next two lines to print cell index
				print (data_in.cell_value(rowx = i, colx = CAP_MAH_COL), 
					data_in.cell_value(rowx = i, colx = CAP_PWR_COL))
				# print (data_in.cell_value(rowx = i, colx = INDEX_COL), 
				# 	data_in.cell_value(rowx = i, colx = CAP_MAH_COL), 
				# 	data_in.cell_value(rowx = i, colx = CAP_PWR_COL))
			i +=1

		print('________________________')




		# ###############
		# # Part 2
		# # Find the voltage drops as discharge starts
		# ###############

		# Initial values do not matter as long as not null
		cur_cell = data_in.cell(rowx = 0, colx = 0)
		last_cell = data_in.cell(rowx = 0, colx = 0)
		last_jump = 1
		cur_jump = 1

		# # Automatically subtracts .001 for rounding purposes 
		for num in range(len(v_target)):
			v_target[num] = v_target[num] - 0.001

		# # Generic counters for looping
		i = 2
		j = 0

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
			if data_in.cell_value(rowx = i, colx = VOL_COL) < CUTOFF:
				print('______________________')
				new_dataset = True
				j = 0
			# Look at current, every falling edge, record position, check values
			if last_cell.value == 0 and cur_cell.value < DISCHARGE_CUR + 1:
				last_jump = cur_jump
				cur_jump = i
				if new_dataset == True:
					print('Temperature = ', round(data_in.cell_value(rowx = cur_jump - 1, 
						colx = TEMP_COL)))
					print('Voltage at Start')
					# Toggle next four lines to print cell index
					print (data_in.cell_value(rowx = cur_jump - 1, colx = VOL_COL))
					print (data_in.cell_value(rowx = cur_jump, colx = VOL_COL))
					#print (data_in.cell_value(rowx = cur_jump - 1, colx = INDEX_COL), 
						#data_in.cell_value(rowx = cur_jump - 1, colx = VOL_COL))
					#print (data_in.cell_value(rowx = cur_jump, colx = INDEX_COL), 
						#data_in.cell_value(rowx = cur_jump, colx = VOL_COL))
					new_dataset = False

				elif (data_in.cell_value(rowx = cur_jump - 1,  colx = VOL_COL) < 
				v_target[j]	< data_in.cell_value(rowx = last_jump - 1,  colx = VOL_COL)):
					print('Voltage = ', v_target[j] + .001)
					# Toggle next four lines to print cell index
					print (data_in.cell_value(rowx = last_jump - 1, colx = VOL_COL))
					print (data_in.cell_value(rowx = last_jump, colx = VOL_COL))			
					#print (data_in.cell_value(rowx = last_jump - 1, colx = INDEX_COL), 
						#data_in.cell_value(rowx = last_jump - 1, colx = VOL_COL))
					#print (data_in.cell_value(rowx = last_jump, colx = INDEX_COL), 
						#data__valuein.cell(rowx = last_jump, colx = VOL_COL))
					j += 1
			i += 1
				
#get_data(filelist)

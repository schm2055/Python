from openpyxl import Workbook, load_workbook
import os, ctypes, getpass, easygui

print'---------------------------------------------------------------------------'
print '                      COMPARE LISTS APPLICATION'
print'---------------------------------------------------------------------------'
print
#Select D42 Document
while True:
	ctypes.windll.user32.MessageBoxA(0, 'Open the D42 Device List', 'Open', 1)
	devices_d = easygui.fileopenbox()
	try:
		devices_wb = load_workbook(filename = devices_d, use_iterators = True)
		break
	except:
		ctypes.windll.user32.MessageBoxA(0, 'Please enter a valid name', 'Import Error!', 1)
		continue
active_sheet_1 = devices_wb.get_sheet_names()[0]
working_sheet_1 = devices_wb.get_sheet_by_name(active_sheet_1) 

#Select Linear Document
while True:
	ctypes.windll.user32.MessageBoxA(0, 'Open the Linear Device List', 'Open', 1)
	devices_l = easygui.fileopenbox()
	try:
		linear_wb = load_workbook(filename = devices_l, use_iterators = True)
		break
	except:
		ctypes.windll.user32.MessageBoxA(0, 'Please enter a valid name', 'Import Error!', 1)
		continue
active_sheet_2 = linear_wb.get_sheet_names()[0]
working_sheet_2 = linear_wb.get_sheet_by_name(active_sheet_2) 

#Create txt document to write results

results_file = 'Comparison_Results.txt'
results_fhand = open(results_file, 'w')

#Run the list comparisons and write the results to the txt file
print 'Checking for devices found in the Linear that were not found in D42...'
results_fhand.write('Checking for devices found in the Linear that were not found in D42...\n\n')

check_list_1 = list()
for row in working_sheet_1:
	for cell in row:
		check_list_1.append(cell.value.upper())

count_1 = 0		
for row in working_sheet_2:
	for cell in row:
		if cell.value.upper() not in check_list_1:
			print cell.value
			results_fhand.write(cell.value + '\n')
			count_1 += 1
		else:
			continue

print 'Number of Devices from the Linear that are not in D42: ', count_1
results_fhand.write('Number of Devices from the Linear that are not in D42: ' + str(count_1) + '\n\n')
print

print 'Checking for devices found in D42 that are not on the Linear Inventory...'
results_fhand.write('Checking for devices found in D42 that are not on the Linear Inventory...\n\n')

check_list_2 = list()
for row in working_sheet_2:
	for cell in row:
		check_list_2.append(cell.value.upper())

count_2 = 0		
for row in working_sheet_1:
	for cell in row:
		if cell.value.upper() not in check_list_2:
			print cell.value
			results_fhand.write(cell.value + '\n')
			count_2 += 1
		else:
			continue			
print 'Number of devices from D42 that are not on the Linear Inventory: ', count_2
results_fhand.write('Number of Devices from the Linear that are not in D42: ' + str(count_2) + '\n\n')
print
print 'Application Complete!'

# -*- coding: utf-8 -*-
"""
Datalog extraction tool using VISA for Agilent 34410A Digital Multimeter
Matthew Sharpe 10-12-18

"""
import sys
import xlsxwriter
import visa
rm = visa.ResourceManager()
inst = rm.list_resources()
picked = 'x'
clear = ''
deviceDict = {}
#build device dictionary
for item in inst:
    deviceDict[inst.index(item) + 1] = item
#user instructions
print('')
print('This program extracts Datalog values in NVMEM via USB from Agilent 34410A DMM.\n')
print('An Excel file will be created in the same directory as this script.\n')
print('Please complete Datalogging manually before running this program.')
print('')

#printing Instrument List
print('Connected devices: ')
for item in deviceDict:
    print(str(item) + ': ' + deviceDict[item])
#pick device loop
while picked not in deviceDict:
    #get user input for device
    try :
        print('')
        picked = int(input('Choose an instrument '))
    except ValueError:
        print("Invalid input! please select Device Number i.e. 1, 2, 3...)\n")
        for item in deviceDict:
            print(str(item) + ': ' + deviceDict[item])
myDevice = deviceDict[picked]
print('')    
print('...reading NVMEM...\n')
#open resource and check for NVMEM data in DMM
try:
    meter = rm.open_resource(myDevice)
    data = meter.query('DATA:DATA? NVMEM')
    points = meter.query('DATA:POINTS? NVMEM')
    lp = points.split('\n')
    print('')
    print(lp[0] + ' points in NVMEM\n')
    if data == '\n':
        print('No data in NVMEM! Check DMM Datalog.')
        sys.exit()
except:
    t = input('Program will exit.')
    sys.exit()
print('')
        
#format data from DMM
l = data.split(',')
f = [float(i) for i in l]

# Create a workbook and add a worksheet.
print('...writing to Excel...\n')
filename = input('Choose a filename: ')
workbook = xlsxwriter.Workbook(filename + '.xlsx')
worksheet = workbook.add_worksheet()
# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0
# Iterate over the data and write it out row by row.
for item in (f):
    worksheet.write(row, col,     item)
    row += 1

workbook.close()

print('Excel workbook created!\n')
try:
    clear = input('Clear NVMEM data? (y/n)')
    if clear == 'y':
        meter.write('DATA:DEL NVMEM')
    else:
        meter.close()
except:
    meter.close()
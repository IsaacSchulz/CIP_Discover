
# Created by Isaac Schulz
# November 27 2024
# Motion Controls Robotics

from xmlrpc.client import DateTime
from pycomm3 import LogixDriver
from pycomm3 import CIPDriver
from pycomm3 import Struct, DINT, STRING, REAL
import xlsxwriter
import time



# Setup the excel file with time stamp
workbook = xlsxwriter.Workbook('CIP_IP_Discover ' + time.asctime(time.localtime()).replace(':', '_') + '.xlsx')
worksheet = workbook.add_worksheet()

# Establish the bold format
bold = workbook.add_format({'bold':True})

# Set all the headings and make them bold
worksheet.write('A1', 'IP ADDRESS', bold)
worksheet.write('B1', 'DEVICE NAME', bold)
worksheet.write('C1', 'VENDOR', bold)
worksheet.write('D1', 'PRODUCT TYPE', bold)
worksheet.write('E1', 'PRODUCT CODE', bold)
worksheet.write('F1', 'FIRMWARE', bold)
worksheet.write('G1', 'SERIAL NUM', bold)
worksheet.write('H1', 'PRODUCT NAME', bold)

# Set up the sorting columns in the workbook
worksheet.set_column(0, 7, 32)

worksheet.autofilter(0, 0, 500, 7)


# Make an empty list to hold all serial numbers written, *1000 is a workaround so I'm just leaving it
serialList = [0]*1000


# Set initial variables for the loop
i = 2
l = 0
n = 0
maxQuery = 10


while n < 10:
    if n > 0:
        print('')
        print('')
        print('Broadcast Query:' + str(n+1) + '/' + str(maxQuery))
        print('---------------------------')
        print('')
    discovered_list = CIPDriver.discover()
    
    for item in discovered_list:
        if item['serial'] not in serialList:

            print(item['ip_address'])
            worksheet.write('A'+str(i), item['ip_address'])

            print(item['vendor'])
            worksheet.write('C'+str(i), item['vendor'])

            print(item['product_type'])
            worksheet.write('D'+str(i), item['product_type'])
    
            print(item['product_code'])
            worksheet.write('E'+str(i), item['product_code'])

            # Firmware is a special item. Each list item in the discovery is a dictionary item EXCEPT for firmware. It is
            # another list of dictionary within that larger list of dictionary for Major and Minor revision. So
            # we have to gather both of them and combine.

            firmware_dict = item['revision']
    
            firmware_string = str(firmware_dict['major']) + '.' + str(firmware_dict['minor'])

            # Print the items to the console to show it working in real time

            print(firmware_string)
            worksheet.write('F'+str(i), firmware_string)

            print(item['serial'])
            worksheet.write('G'+str(i), item['serial'])

            serialList[l] = item['serial']

            print(item['product_name'])
            worksheet.write('H'+str(i), item['product_name'])

            i = i + 1
            l = l + 1
            print(' ')
    n = n + 1

    CIPDriver.close

    
# Allow the user to see the console before closing the window and ending program
print('Revision November 27 2024')

workbook.close()

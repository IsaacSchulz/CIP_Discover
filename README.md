# CIP_Discover
Discovering all CIP communicating devices on a network

This uses Pycomm3 at the heart. https://docs.pycomm3.dev/en/latest/

This scans the network your windows PC is on and returns an excel sheet listing details of each device:

IP Address,
Vendor,
Product Type,
Product Code,
Firmware,
Serial Number,
Product Name

and creates an extra column titles 'Device Name' for the user to enter their own nicknames for the devices.

The Discover function in Pycomm3 is not perfect. Running it once can miss some items that don't respond right away. So, I made it so the script runs the discover query 10 times and adds any new items to the list before writing to the excel document.

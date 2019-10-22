#! python3

import re, pyperclip, openpyxl, os, pprint

# Create diretory if it doesn't exists
if os.path.exists('c:\\TIA Tags') == False:
    os.makedirs('c:\\TIA Tags')

# Change current work directory
os.chdir('c:\\TIA Tags')

# Create a list for spreadsheet header
headerString = 'Name	Path	Connection	PLC tag	DataType	HMI DataType	Length	Coding	Access Method	Address	Start value	Quality Code	Persistency	Substitute value	Tag value [pt-BR]	Update Mode	Comment [pt-BR]	Limit Upper 2 Type	Limit Upper 2	Limit Lower 2 Type	Limit Lower 2	Linear scaling	End value PLC	Start value PLC	End value HMI	Start value HMI	Synchronization'
header = headerString.split('\t')

# Create a spreadsheet with list header
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = 'Hmi Tags'
for i in header:
    sheet.cell(row=1, column=header.index(i) + 1).value = i

# Create a regex for information
MotorRegex = re.compile(r'''
.*?\d\..*?                    #getTableNumber
Tag.*?(\w+).*?                #getTagName
Input\t(\d{3}\.\d:I).*?       #getInputNumberSymbol
Output\t(\d{3}\.\d:O).*?      #getOutputNumberSymbol
Panel\t(\w+\d+).*?            #getPanelName
TimeRunning.*?\t(\d+).*?      #getTimeRunning  
Description\t(\w+\s+\w+).*?   #getDescription
''', re.VERBOSE | re.DOTALL)

# Get the text off the clipboard
text = pyperclip.paste()

# Extract the storage from this text
extractedMotors = MotorRegex.findall(text)

print(extractedMotors)

# excel line counter
j = 2

PLCName = 'PLC_1'

for motor in extractedMotors:
    # tag Input
    sheet.cell(row=j, column=1).value = PLCName + '_' + motor[0] + '_Input'
    sheet.cell(row=j, column=2).value = 'Project\\Motores\\' + motor[0]
    sheet.cell(row=j, column=3).value = PLCName
    sheet.cell(row=j, column=4).value = '"' + PLCName + '_' + motor[0] + '.Input'
    sheet.cell(row=j, column=5).value = 'DInt'
    sheet.cell(row=j, column=6).value = 'DInt'
    sheet.cell(row=j, column=7).value = 4
    sheet.cell(row=j, column=8).value = 'Binary'
    sheet.cell(row=j, column=9).value = 'Symbolic access'
    sheet.cell(row=j, column=10).value = '<No Value>'
    sheet.cell(row=j, column=11).value = '<No Value>'
    sheet.cell(row=j, column=12).value = 'False'
    sheet.cell(row=j, column=13).value = 'False'
    sheet.cell(row=j, column=14).value = '<No Value>'
    sheet.cell(row=j, column=15).value = '<No Value>'
    sheet.cell(row=j, column=16).value = 'Client/Server wide'
    sheet.cell(row=j, column=17).value = '<No Value>'
    sheet.cell(row=j, column=18).value = 'None'
    sheet.cell(row=j, column=19).value = '<No Value>'
    sheet.cell(row=j, column=20).value = 'None'
    sheet.cell(row=j, column=21).value = '<No Value>'
    sheet.cell(row=j, column=22).value = 'False'
    sheet.cell(row=j, column=23).value = 10
    sheet.cell(row=j, column=24).value = 0
    sheet.cell(row=j, column=25).value = 100
    sheet.cell(row=j, column=26).value = 0
    sheet.cell(row=j, column=27).value = 'False'

    # tag Output
    sheet.cell(row=j+1, column=1).value = PLCName + '_' + motor[0] + '_Output'
    sheet.cell(row=j+1, column=2).value = 'Project\\Motores\\' + motor[0]
    sheet.cell(row=j+1, column=3).value = PLCName
    sheet.cell(row=j+1, column=4).value = '"' + PLCName + '_' + motor[0] + '.Output'
    sheet.cell(row=j+1, column=5).value = 'DInt'
    sheet.cell(row=j+1, column=6).value = 'DInt'
    sheet.cell(row=j+1, column=7).value = 4
    sheet.cell(row=j+1, column=8).value = 'Binary'
    sheet.cell(row=j+1, column=9).value = 'Symbolic access'
    sheet.cell(row=j+1, column=10).value = '<No Value>'
    sheet.cell(row=j+1, column=11).value = '<No Value>'
    sheet.cell(row=j+1, column=12).value = 'False'
    sheet.cell(row=j+1, column=13).value = 'False'
    sheet.cell(row=j+1, column=14).value = '<No Value>'
    sheet.cell(row=j+1, column=15).value = '<No Value>'
    sheet.cell(row=j+1, column=16).value = 'Client/Server wide'
    sheet.cell(row=j+1, column=17).value = '<No Value>'
    sheet.cell(row=j+1, column=18).value = 'None'
    sheet.cell(row=j+1, column=19).value = '<No Value>'
    sheet.cell(row=j+1, column=20).value = 'None'
    sheet.cell(row=j+1, column=21).value = '<No Value>'
    sheet.cell(row=j+1, column=22).value = 'False'
    sheet.cell(row=j+1, column=23).value = 10
    sheet.cell(row=j+1, column=24).value = 0
    sheet.cell(row=j+1, column=25).value = 100
    sheet.cell(row=j+1, column=26).value = 0
    sheet.cell(row=j+1, column=27).value = 'False'

    # tag Panel
    sheet.cell(row=j+2, column=1).value = PLCName + '_' + motor[0] + '_Panel' 
    sheet.cell(row=j+2, column=2).value = 'Project\\Motores\\' + motor[0]
    sheet.cell(row=j+2, column=3).value = PLCName
    sheet.cell(row=j+2, column=4).value = '<No Value>'
    sheet.cell(row=j+2, column=5).value = 'TextRef'
    sheet.cell(row=j+2, column=6).value = 'TextRef'
    sheet.cell(row=j+2, column=7).value = 4
    sheet.cell(row=j+2, column=8).value = 'Binary'
    sheet.cell(row=j+2, column=9).value = '<No Value>'
    sheet.cell(row=j+2, column=10).value = '<No Value>'
    sheet.cell(row=j+2, column=11).value = '<No Value>'
    sheet.cell(row=j+2, column=12).value = 'False'
    sheet.cell(row=j+2, column=13).value = 'False'
    sheet.cell(row=j+2, column=14).value = '<No Value>'
    sheet.cell(row=j+2, column=15).value = motor[3]
    sheet.cell(row=j+2, column=16).value = 'Client/Server wide'
    sheet.cell(row=j+2, column=17).value = '<No Value>'
    sheet.cell(row=j+2, column=18).value = 'None'
    sheet.cell(row=j+2, column=19).value = '<No Value>'
    sheet.cell(row=j+2, column=20).value = 'None'
    sheet.cell(row=j+2, column=21).value = '<No Value>'
    sheet.cell(row=j+2, column=22).value = 'False'
    sheet.cell(row=j+2, column=23).value = 10
    sheet.cell(row=j+2, column=24).value = 0
    sheet.cell(row=j+2, column=25).value = 100
    sheet.cell(row=j+2, column=26).value = 0
    sheet.cell(row=j+2, column=27).value = 'False'

    # tag Output
    sheet.cell(row=j+3, column=1).value = PLCName + '_' + motor[0] + '_TimeRunning'
    sheet.cell(row=j+3, column=2).value = 'Project\\Motores\\' + motor[0]
    sheet.cell(row=j+3, column=3).value = PLCName
    sheet.cell(row=j+3, column=4).value = '"' + PLCName + '_' + motor[0] + '.TimeRunning'
    sheet.cell(row=j+3, column=5).value = 'DInt'
    sheet.cell(row=j+3, column=6).value = 'DInt'
    sheet.cell(row=j+3, column=7).value = 4
    sheet.cell(row=j+3, column=8).value = 'Binary'
    sheet.cell(row=j+3, column=9).value = 'Symbolic access'
    sheet.cell(row=j+3, column=10).value = '<No Value>'
    sheet.cell(row=j+3, column=11).value = '<No Value>'
    sheet.cell(row=j+3, column=12).value = 'False'
    sheet.cell(row=j+3, column=13).value = 'False'
    sheet.cell(row=j+3, column=14).value = '<No Value>'
    sheet.cell(row=j+3, column=15).value = motor[4]
    sheet.cell(row=j+3, column=16).value = 'Client/Server wide'
    sheet.cell(row=j+3, column=17).value = '<No Value>'
    sheet.cell(row=j+3, column=18).value = 'None'
    sheet.cell(row=j+3, column=19).value = '<No Value>'
    sheet.cell(row=j+3, column=20).value = 'None'
    sheet.cell(row=j+3, column=21).value = '<No Value>'
    sheet.cell(row=j+3, column=22).value = 'False'
    sheet.cell(row=j+3, column=23).value = 10
    sheet.cell(row=j+3, column=24).value = 0
    sheet.cell(row=j+3, column=25).value = 100
    sheet.cell(row=j+3, column=26).value = 0
    sheet.cell(row=j+3, column=27).value = 'False'

    # tag Description
    sheet.cell(row=j+4, column=1).value = PLCName + '_' + motor[0] + '_Description' 
    sheet.cell(row=j+4, column=2).value = 'Project\\Motores\\' + motor[0]
    sheet.cell(row=j+4, column=3).value = PLCName
    sheet.cell(row=j+4, column=4).value = '<No Value>'
    sheet.cell(row=j+4, column=5).value = 'TextRef'
    sheet.cell(row=j+4, column=6).value = 'TextRef'
    sheet.cell(row=j+4, column=7).value = 4
    sheet.cell(row=j+4, column=8).value = 'Binary'
    sheet.cell(row=j+4, column=9).value = '<No Value>'
    sheet.cell(row=j+4, column=10).value = '<No Value>'
    sheet.cell(row=j+4, column=11).value = '<No Value>'
    sheet.cell(row=j+4, column=12).value = 'False'
    sheet.cell(row=j+4, column=13).value = 'False'
    sheet.cell(row=j+4, column=14).value = '<No Value>'
    sheet.cell(row=j+4, column=15).value = motor[5]
    sheet.cell(row=j+4, column=16).value = 'Client/Server wide'
    sheet.cell(row=j+4, column=17).value = '<No Value>'
    sheet.cell(row=j+4, column=18).value = 'None'
    sheet.cell(row=j+4, column=19).value = '<No Value>'
    sheet.cell(row=j+4, column=20).value = 'None'
    sheet.cell(row=j+4, column=21).value = '<No Value>'
    sheet.cell(row=j+4, column=22).value = 'False'
    sheet.cell(row=j+4, column=23).value = 10
    sheet.cell(row=j+4, column=24).value = 0
    sheet.cell(row=j+4, column=25).value = 100
    sheet.cell(row=j+4, column=26).value = 0
    sheet.cell(row=j+4, column=27).value = 'False'
    j += 5    

# Save the spreadsheet on path
wb.save('MotorTags.xlsx')        

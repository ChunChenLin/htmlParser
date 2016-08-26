from bs4 import BeautifulSoup
import collections
import xlsxwriter

file1 = raw_input("(1)Please enter the file name(no need to include .html)")+".html"
file2 = raw_input("(2)Please enter the file name(no need to include .html)")+".html"

res1 = open(file1)
soup1 = BeautifulSoup(res1)
res2 = open(file2)
soup2 = BeautifulSoup(res2)

collect = collections.defaultdict(list) #dict type

Basic = ['SR Number',
         'Product Name',
         'Version Type',
         'Trial Days',
         'Company Name',
         'CD Key',
         'Bit Version',
         'Installation Type',
         'Delivery Format',
         'Language',
         'Dolby Share Compoent',
         'Hardware Lock in setup process',
         'Accessaries',
         'BU',
         'Installation Path',
         'Desktop Shortcut',
         'StartMenu',
         'Add/Remove Program',
         'Limit SR',
         'Customer ini',
         'Send Installer Logs',
         'Inherit Previous CDKey to Upgrade',
         'CDKey Online Activation',
         'Package Type',
         'Dolby Codec Program License Check (Windows 8)',
         'Working Directory',
         'Interop with PowerProducer',
         'Point of Use (POU)',
         'Output Channels',
         'SmartSound',
         'Default SubSR (for no command line case)']

Royalty = ['H.264/MPEG-4 AVC to AT&T',
           'H.264/MPEG-4 AVC to MPEG-LA',
           'MPEG-2/4 AAC to VIA Licensing',
           'H.265/MPEG HEVC to HEVC Advance LLC',
           'H.265/MPEG HEVC to MPEG-LA',
           'MPEG-2 Video']

for i in Basic:
    collect[i].append('NA')
    collect[i].append('NA')
for i in Royalty:
    collect[i].append('NA')
    collect[i].append('NA')

number1 = int(raw_input("(3)Please enter the Number of Basic you want to compare:"))
number2 = int(raw_input("(4)Please enter the Number of Basic you want to compare:"))
output = raw_input("(5)Please enter the file name you want to export:")

start1 = 0
end1 = 0
count = 0
n = 0
for i in soup1.select('td'):
    end1 = n
    if "Basic" in i.text:
        count += 1
        if count == number1:
            start1 = n
        if count == number1 + 1:
            break
    n += 1

#print count,start1, end1

start2 = 0
end2 = 2000
count = 0
n = 0
for i in soup2.select('td'):
    end2 = n
    if "Basic" in i.text:
        count += 1
        if count == number2:
            start2 = n
        if count == number2 + 1:
            break
    n += 1       

for k, v in collect.items():
    for i in range(start1,end1):
        if k in soup1.select('td')[i].text:
            collect[k][0] = soup1.select('td')[i+1].text
            break
    for j in range(start2,end2):
        if k in soup2.select('td')[j].text:
            collect[k][1] = soup2.select('td')[j+1].text
            break

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook(output+'.xlsx')
worksheet = workbook.add_worksheet()
# Add a format to use to highlight cells.
format = workbook.add_format({'bold': True, 'font_color': 'red'})
diff = workbook.add_format({'bold': True, 'font_color': 'blue'})
row = 1
col = 0
worksheet.write(0, col, "Basic&Royalty", format)
#worksheet.write(0, col + 1, collect['SR Number'][0], format)
#worksheet.write(0, col + 2, collect['SR Number'][1], format)
for k, v in collect.items():
    #print k, v[0], v[1]
    worksheet.write(row, col, k)
    if v[0] != v[1]:
        worksheet.write(row, col + 1, v[0], diff)
        worksheet.write(row, col + 2, v[1], diff)
    else:
        worksheet.write(row, col + 1, v[0])
        worksheet.write(row, col + 2, v[1])
    row += 1

workbook.close()
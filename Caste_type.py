from pathlib import Path
import xlrd
import xlwt
from xlwt import Workbook
path = Path("copy.xlsx")  #path of excel file you want to open
wb = xlrd.open_workbook(path)
w = Workbook()
sheet = wb.sheet_by_index(0)
sheet1 = w.add_sheet('Sheet 1')

no_of_columns = 650

def writeColumns():
    sheet1.write(0, 4, 'Type')

def writeName():
    for i in range(no_of_columns):
        for j in range(4):
            name = sheet.cell_value(i,j)
            sheet1.write(i, j, name)

def writetype():
    count = 1
    for i in range(1, no_of_columns):
        for j in range(1,2):
            value = sheet.cell_value(i,j)
            if "JAIN" in f'{value}':
                sheet1.write(count, 4, 1)
                count+=1
                print(sheet.cell_value(i,j))
            elif "AGRAWAL" in f'{value}' or "AGGARWAL" in f'{value}' or "GUPTA" in f'{value}' or "GARG" in f'{value}' or "VAISH" in f'{value}' or "SINGHLA" in f'{value}' or "JAISWAL" in f'{value}' or "BERNWAL" in f'{value}':
                sheet1.write(count, 4, 2)
                count+=1
                print(sheet.cell_value(i,j))
            elif "BALI" in f'{value}' or "MOHIYAL" in f'{value}' or "BASI" in f'{value}' or "BHARDWAJ" in f'{value}' or "GAUR" in f'{value}' or "SHARMA" in f'{value}' or "MISHRA" in f'{value}' or "SHUKLA" in f'{value}' or "TRIPATHI" in f'{value}' or "PANDEY" in f'{value}' or "PATHAK" in f'{value}' or "DIXIT" in f'{value}' or "CHATURVEDI" in f'{value}' or "TRIVEDI" in f'{value}' or "DUBEY" in f'{value}' or "VASHISHTH" in f'{value}' or "JHA"in f'{value}':
                sheet1.write(count, 4, 3)
                count+=1
                print(sheet.cell_value(i, j))
            elif "SINGH" in f'{value}' or "YADAV" in f'{value}' or "SHRIVASTAVA" in f'{value}' or "SINHA"in f'{value}' or "VERMA" in f'{value}' or "PAL" in f'{value}' or "MAURYA" in f'{value}' or "KUSHWAHA" in f'{value}' or "RAJBHAR" in f'{value}' or "PASWAN" in f'{value}' or "CHAUHAN" in f'{value}' or "NISHAD" in f'{value}' or "PATEL" in f'{value}':
                sheet1.write(count, 4, 4)
                count+=1
                print(sheet.cell_value(i, j))
            elif "JOSEPH" in f'{value}' or "JOHNSON" in f'{value}' or "JOHN"in f'{value}':
                sheet1.write(count, 4, 5)
                count += 1
                print(sheet.cell_value(i, j))
            elif "MACSOOD" in f'{value}' or "ALI" in f'{value}' or f'KHAN'in f'{value}' or "ISLAM" in f'{value}' or "FAZAL" in f'{value}' or "MALLIK" in f'{value}' or "BEGUM" in f'{value}' or "HAMID" in f'{value}' or "SHEIKH" in f'{value}' or "SARFRAZ" in f'{value}' or "SHERVANI" in f'{value}':
                sheet1.write(count, 4, 6)
                count += 1
                print(sheet.cell_value(i, j))
            elif "KAUR" in f'{value}' or  "KUCKREJA" in f'{value}' or f"KHANNA" in f'{value}' or "CHAWLA" in f'{value}' or "KUKREJA" in f'{value}' or "SETH" in f'{value}' or "KHATRI" in f'{value}' or "CHOPRA" in f'{value}' or f"SETHI" in f'{value}':
                sheet1.write(count, 4, 7)
                count += 1
                print(sheet.cell_value(i, j))
            else:
                sheet1.write(count, 4, 8)
                count+=1
                print(sheet.cell_value(i,j))

    print(count)

writeColumns()
writeName()
writetype()

w.save('copyfinal.xls') #Name of the new excel file

from pyautocad import Autocad, APoint
import openpyxl
import math

path = 'ALl Co-ordinates.xlsx'
acad = Autocad()
acad.prompt("Hello, Autocad from Python\n")
print(acad.doc.Name)


def read_from_column(sheet, column, start_row, no_of_data):
    data = []
    for i in range(0,no_of_data):
        data.append(sheet.cell(start_row+i,column).value)
    return data
wb = openpyxl.load_workbook(path)
sheet = wb.active
print(sheet.title)

Major_east = read_from_column(sheet, 3, 4, 25)
Major_north = read_from_column(sheet, 4, 4, 25)
Minor_east = read_from_column(sheet, 3, 29,8)
Minor_north = read_from_column(sheet, 4, 29,8)
IP_east = read_from_column(sheet, 3, 37, 18)
IP_north = read_from_column(sheet, 4, 37, 18)
BC_east = read_from_column(sheet, 8, 4, 12)
BC_north = read_from_column(sheet, 9, 4, 12)
EC_east = read_from_column(sheet, 13, 4, 12)
EC_north = read_from_column(sheet, 14, 4, 12)
MC_east = read_from_column(sheet, 19, 4, 12)
MC_north = read_from_column(sheet, 20, 4, 12)
Radius = read_from_column(sheet, 16, 4, 13)
print('DATA READ')
for i in range(len(Major_east)-1):
    p1 = APoint(Major_east[i], Major_north[i])
    p2 = APoint(Major_east[i+1], Major_north[i+1])
    acad.model.AddLine(p1, p2)
    acad.model.AddCircle(p1,5)
    acad.model.AddCircle(p1,3)
print('MAJOR TRAVERSE')
for i in range(len(Minor_east)-1):
    p1 = APoint(Minor_east[i], Minor_north[i])
    p2 = APoint(Minor_east[i+1], Minor_north[i+1])
    acad.model.AddLine(p1, p2)
    if(i==0):
        continue
    acad.model.AddCircle(p1,3)
print('MINOR TRAVERSE')
for i in range(len(IP_east)-1):
    p1 = APoint(IP_east[i], IP_north[i])
    p2 = APoint(IP_east[i+1], IP_north[i+1])
    acad.model.AddLine(p1, p2)
print('ROAD CENTER-LINE')
#for i in range(len(BC_east)):
#    p1 = (BC_east[i],BC_north[i])
#    p2 = (MC_east[i],MC_north[i])
#    p3 = (EC_east[i],EC_north[i])
#    r = Radius[i]
#    a1 = 2*(p1[0]-p2[0])
#    b1 = 2*(p1[1]-p2[1])
#    c1 = (p1[0]**2)-(p2[0]**2)+(p1[1]**2)-(p2[1]**2)
#    a2 = 2*(p1[0]-p3[0])
#    b2 = 2*(p1[1]-p3[1])
#    c2 = (p1[0]**2)-(p3[0]**2)+(p1[1]**2)-(p3[1]**2)
#    k = (c2-a2*c1/a1)/(b2-a2*b1/a1)
#    h = (c1-b1*k)/a1
#    acad.model.AddCircle(APoint(h,k),r)
#print('CURVES')




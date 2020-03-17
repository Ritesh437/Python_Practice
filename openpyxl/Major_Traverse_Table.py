import openpyxl
import progressbar
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Font, PatternFill

wb = openpyxl.Workbook()
sheet1 = wb.active
sheet1.sheet_view.showGridLines = False
sheet1.title = 'Major'

t_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
th_border = Border(left=Side(style='thick'), right=Side(style='thick'), top=Side(style='thick'), bottom=Side(style='thick'))
bold_grey = Font(bold=True, color='a9a9a9')
big_Text = Font(size=20)
color_red = PatternFill(start_color='ffb0b0',fill_type='solid')
color_green = PatternFill(start_color='b1ffb0',fill_type='solid')


n_major = int(input('Input No Of Stations: '))

# n_minor = int(input('Input No Of Minor Stations: '))

# start = int(input('Starting Station(integer value): '))

# end = int(input('Ending Station(integer value): '))
# n_major = 11
# n_minor = 4
# start = 1
# end = 8

def bg_color_range(sheet, cell_range, pattern):
    start_cell, end_cell = cell_range.split(':')
    start_coord = openpyxl.utils.cell.coordinate_from_string(start_cell)
    start_row = start_coord[1]
    start_col = openpyxl.utils.cell.column_index_from_string(start_coord[0])
    end_coord = openpyxl.utils.cell.coordinate_from_string(end_cell)
    end_row = end_coord[1]
    end_col = openpyxl.utils.cell.column_index_from_string(end_coord[0])

    for row in range(start_row, end_row + 1):
        for col_idx in range(start_col, end_col + 1):
            sheet.cell(row,col_idx).fill = pattern

def style_range(sheet, cell_range, style):

    start_cell, end_cell = cell_range.split(':')
    start_coord = openpyxl.utils.cell.coordinate_from_string(start_cell)
    start_row = start_coord[1]
    start_col = openpyxl.utils.cell.column_index_from_string(start_coord[0])
    end_coord = openpyxl.utils.cell.coordinate_from_string(end_cell)
    end_row = end_coord[1]
    end_col = openpyxl.utils.cell.column_index_from_string(end_coord[0])

    for row in range(start_row, end_row + 1):
        for col_idx in range(start_col, end_col + 1):
            sheet.cell(row,col_idx).font = style
 
def border_range(sheet, cell_range, border):

    start_cell, end_cell = cell_range.split(':')
    start_coord = openpyxl.utils.cell.coordinate_from_string(start_cell)
    start_row = start_coord[1]
    start_col = openpyxl.utils.cell.column_index_from_string(start_coord[0])
    end_coord = openpyxl.utils.cell.coordinate_from_string(end_cell)
    end_row = end_coord[1]
    end_col = openpyxl.utils.cell.column_index_from_string(end_coord[0])

    for row in range(start_row, end_row + 1):
        for col_idx in range(start_col, end_col + 1):
            sheet.cell(row,col_idx).border = border


style_range(sheet1,f'A1:A{sheet1.max_column}',big_Text)


sheet1.cell(1,1).value = 'Closed Traverse Calculation'
sheet1.merge_cells('A1:C1')
sheet1.cell(1,4).value = 'No of Stations : '
sheet1.merge_cells('D1:E1')
sheet1.cell(1,6).value = n_major
sheet1.cell(2,1).value = 'Station'
sheet1.merge_cells('A2:A3')
sheet1.cell(2,2).value = 'Leg'
sheet1.merge_cells('B2:B3')
sheet1.cell(2,3).value = 'Leg Length'
sheet1.merge_cells('C2:C3')
sheet1.cell(2,4).value = 'Hz Angle'
sheet1.merge_cells('D2:F2')
sheet1.cell(3,4).value = 'D'
sheet1.cell(3,5).value = 'M'
sheet1.cell(3,6).value = 'S'
sheet1.cell(2,7).value = 'Correction'
sheet1.merge_cells('G2:I2')
sheet1.cell(3,7).value = 'D'
sheet1.cell(3,8).value = 'M'
sheet1.cell(3,9).value = 'S'
sheet1.cell(2,10).value = 'Corrected Angles'
sheet1.merge_cells('J2:L2')
sheet1.cell(3,10).value = 'D'
sheet1.cell(3,11).value = 'M'
sheet1.cell(3,12).value = 'S'
sheet1.cell(2,13).value = 'Bearing'
sheet1.merge_cells('M2:O2')
sheet1.cell(3,13).value = 'D'
sheet1.cell(3,14).value = 'M'
sheet1.cell(3,15).value = 'S'
sheet1.cell(2,16).value = 'Consecutive Coordinates'
sheet1.merge_cells('P2:Q2')
sheet1.cell(3,16).value = 'Dep'
sheet1.cell(3,17).value = 'Lat'
sheet1.cell(2,18).value = 'Correction'
sheet1.merge_cells('R2:S2')
sheet1.cell(3,18).value = 'Dep'
sheet1.cell(3,19).value = 'Lat'
sheet1.cell(2,20).value = 'Corrected Consecutive Coordinates'
sheet1.merge_cells('T2:U2')
sheet1.cell(3,20).value = 'Dep'
sheet1.cell(3,21).value = 'Lat'
sheet1.cell(2,22).value = 'Independent Coordinates'
sheet1.merge_cells('V2:W2')
sheet1.cell(3,22).value = 'Easting'
sheet1.cell(3,23).value = 'Northing'
sheet1.cell(2,24).value = 'Adjusted Length'
sheet1.merge_cells('X2:X3')
sheet1.cell(2,25).value = 'Adjusted Bearing'
sheet1.merge_cells('Y2:Y3')


for i in range (1,n_major+1):
    sheet1.cell(i+3,1).value = f'M{i}'
    if(i==n_major):
        sheet1.cell(i+3,2).value = f'M{i}M{1}'
        continue

    sheet1.cell(i+3,2).value = f'M{i}M{i+1}'


for i in range(1,sheet1.max_row+1):
    for j in range(1,sheet1.max_column+1):
        sheet1.cell(i,j).alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

style_range(sheet1,'A2:Y3',bold_grey)
style_range(sheet1,f'A4:B{sheet1.max_row}',bold_grey)

border_range(sheet1, f'A2:Y{sheet1.max_row}',t_border)
border_range(sheet1, f'A2:Y3',th_border)

rd = sheet1.row_dimensions[1]
rd.height = 55
rd = sheet1.row_dimensions[2]
rd.height = 45

def set_col_width(sheet,a,b,size):
    for i in range(a,b+1):
        col_str = openpyxl.utils.cell.get_column_letter(i)
        cd = sheet.column_dimensions[col_str]
        cd.width = size

bg_color_range(sheet1,f'C4:F{sheet1.max_row}',color_red)
bg_color_range(sheet1,f'G4:Y{sheet1.max_row}',color_green)
bg_color_range(sheet1,'M4:O4',color_red)
bg_color_range(sheet1,'V4:W4',color_red)


set_col_width(sheet1,4,15,6)
set_col_width(sheet1,22,23,12)

wb.save('Traverse.xlsx')
print('Table File Created >> Traverse.xlsx')
print('Fill The Data in Red Cells, close the file and Continue From Calculate.py')


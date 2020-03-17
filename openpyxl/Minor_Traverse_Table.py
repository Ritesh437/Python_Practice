
sheet2 = wb.create_sheet('Minor')
sheet2.sheet_view.showGridLines = False

style_range(sheet2,f'A1:A{sheet2.max_column}',big_Text)

sheet2.cell(1,1).value = 'Minor Traverse Calculation'
sheet2.merge_cells('A1:C1')
sheet2.cell(1,4).value = 'No of Stations : '
sheet2.merge_cells('D1:E1')
sheet2.cell(1,6).value = n_minor
sheet2.cell(1,7).value = 'Start:'
sheet2.cell(1,8).value = f'M{start}'
sheet2.cell(1,9).value = 'End:'
sheet2.cell(1,10).value = f'M{end}'
sheet2.cell(2,1).value = 'Station'
sheet2.merge_cells('A2:A3')
sheet2.cell(2,2).value = 'Leg'
sheet2.merge_cells('B2:B3')
sheet2.cell(2,3).value = 'Leg Length'
sheet2.merge_cells('C2:C3')
sheet2.cell(2,4).value = 'Hz Angle'
sheet2.merge_cells('D2:F2')
sheet2.cell(3,4).value = 'D'
sheet2.cell(3,5).value = 'M'
sheet2.cell(3,6).value = 'S'
sheet2.cell(2,7).value = 'Bearing'
sheet2.merge_cells('G2:I2')
sheet2.cell(3,7).value = 'D'
sheet2.cell(3,8).value = 'M'
sheet2.cell(3,9).value = 'S'
sheet2.cell(2,10).value = 'Correction'
sheet2.merge_cells('J2:L2')
sheet2.cell(3,10).value = 'D'
sheet2.cell(3,11).value = 'M'
sheet2.cell(3,12).value = 'S'
sheet2.cell(2,13).value = 'Corrected Bearing'
sheet2.merge_cells('M2:O2')
sheet2.cell(3,13).value = 'D'
sheet2.cell(3,14).value = 'M'
sheet2.cell(3,15).value = 'S'
sheet2.cell(2,16).value = 'Consecutive Coordinates'
sheet2.merge_cells('P2:Q2')
sheet2.cell(3,16).value = 'Dep'
sheet2.cell(3,17).value = 'Lat'
sheet2.cell(2,18).value = 'Correction'
sheet2.merge_cells('R2:S2')
sheet2.cell(3,18).value = 'Dep'
sheet2.cell(3,19).value = 'Lat'
sheet2.cell(2,20).value = 'Corrected Consecutive Coordinates'
sheet2.merge_cells('T2:U2')
sheet2.cell(3,20).value = 'Dep'
sheet2.cell(3,21).value = 'Lat'
sheet2.cell(2,22).value = 'Independent Coordinates'
sheet2.merge_cells('V2:W2')
sheet2.cell(3,22).value = 'Easting'
sheet2.cell(3,23).value = 'Northing'
sheet2.cell(2,24).value = 'Adjusted Length'
sheet2.merge_cells('X2:X3')
sheet2.cell(2,25).value = 'Adjusted Bearing'
sheet2.merge_cells('Y2:Y3')

sheet2.cell(4,1).value = f'M{start}'
if (start == 1):
    sheet2.cell(4,2).value = f'M{n_major}M{start}'
else:
    sheet2.cell(4,2).value = f'M{start-1}M{start}'

for i in range(1,n_minor+2):
    sheet2.cell(i+4,1).value = f'm{i}'
    if(i == 1):
        sheet2.cell(i+4,2).value = f'M{start}m{i}'
        continue
    if(i == n_minor):
        sheet2.cell(i+4,2).value = f'm{i-1}M{end}'
        continue

    sheet2.cell(i+4,2).value = f'm{i-1}m{i}'


sheet2.cell(5+n_minor,1).value = f'M{end}'
if (end == 1):
    sheet2.cell(5+n_minor,2).value = f'M{n_major}M{end}'
else:
    sheet2.cell(5+n_minor,2).value = f'M{end-1}M{end}'

set_col_width(sheet2,4,15,6)
set_col_width(sheet2,22,23,12)

bg_color_range(sheet2,f'C4:F{sheet2.max_row}',color_red)
bg_color_range(sheet2,f'G4:Y{sheet2.max_row}',color_green)

rd = sheet2.row_dimensions[1]
rd.height = 55
rd = sheet2.row_dimensions[2]
rd.height = 45

border_range(sheet2, f'A2:Y{sheet2.max_row}',t_border)
border_range(sheet2, f'A2:Y3',th_border)

style_range(sheet2,'A2:Y3',bold_grey)
style_range(sheet2,f'A4:B{sheet2.max_row}',bold_grey)

for i in range(1,sheet2.max_row+1):
    for j in range(1,sheet2.max_column+1):
        sheet2.cell(i,j).alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

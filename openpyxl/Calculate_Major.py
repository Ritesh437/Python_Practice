import openpyxl
import math
import pygame
from pygame.locals import*

path = r'D:\ritisha.xlsx'
wb = openpyxl.load_workbook(path,read_only=False)
sheet1 = wb['Major']

n_major = sheet1.cell(1,6).value

def read_from_column(sheet, column, start_row, no_of_data):
    data = []
    for i in range(0,no_of_data):
        data.append(sheet.cell(start_row+i,column).value)
    return data


leg_lengths = read_from_column(sheet1, 3, 4, n_major)

print(leg_lengths)

Deg = read_from_column(sheet1, 4, 4, n_major)
Min = read_from_column(sheet1, 5, 4, n_major)
Sec = read_from_column(sheet1, 6, 4, n_major)

bearing = [sheet1.cell(4,13).value, sheet1.cell(4,14).value, sheet1.cell(4,15).value]
print(bearing)
easting = sheet1.cell(4,22).value
northing = sheet1.cell(4,23).value
print(northing,'  ',easting)
class AngleDMS:
    Deg = []
    Min = []
    Sec = []

    def __init__(self, Deg, Min, Sec):
        self.Deg = Deg
        self.Min = Min
        self.Sec = Sec

hz_angle = AngleDMS(Deg, Min, Sec)

def calculate_bearing(bea,ang):
    sm = dmsAddition(bea,ang)
    sm_degree = convert_to_deg(sm)
    if(sm_degree>=540):
        sm = dmsSubtraction(sm,[540,0,0])
    else:
        if(sm_degree>=180):
            sm = dmsSubtraction(sm,[180,0,0])
        else:
            sm = dmsAddition(sm,[180,0,0])
    return sm

def round_up(x):
    aX = abs(x)
    return int(math.ceil(aX)*(aX/x)) if x != 0 else 0


def convert_to_dms(Degree):
    Deg = math.trunc(Degree)
    minn = (Degree-Deg)*60
    Min = math.trunc(minn)
    Sec = round((minn-Min)*60, 4)
    return [Deg, Min, Sec]

def convert_to_deg(dms):
    Degree = (dms[0])+(dms[1]/60)+(dms[2]/3600)
    return Degree


def dmsDivision(a,n):
    div = [a[0]/n, a[1]/n, a[2]/n]
    b = (div[0]-math.trunc(div[0]))*60
    div[0] = math.trunc(div[0])
    div[1] = div[1] + b
    c = (div[1]-math.trunc(div[1]))*60
    div[1] = math.trunc(div[1])
    div[2] = div[2] + c
    return div

def dmsSubtraction(a,b):
    c = a[0] + a[1]/60 + a[2]/3600 
    d = b[0] + b[1]/60 + b[2]/3600
    dif = c - d
    ab = convert_to_dms(abs(dif))
    x = ab[0]
    y = ab[1]
    z = ab[2]
    if(dif>0):
        return [x,y,z]
    else:
        if(dif<0):
            return [-x,-y,-z]
        else:
            return [0,0,0]

def dmsAddition(a,b):
    Deg = a[0] + b[0]
    Min = a[1] + b[1]
    Sec = a[2] + b[2]
    if (Sec >= 60):
        Sec -= 60
        Min += 1
    if(Min >= 60):
        Min -= 60
        Deg += 1
    return [Deg, Min, Sec]

def angleCorrection(hz_angle, corr, n):
    Deg = []
    corDeg = []
    Min = []
    corMin = []
    Sec = []
    corSec = []
    check = []
    Degree = []
    for i in range(len(hz_angle.Deg)):
        check.append(convert_to_deg([hz_angle.Deg[i],hz_angle.Min[i],hz_angle.Sec[i]]))
        Degree.append(convert_to_deg([hz_angle.Deg[i],hz_angle.Min[i],hz_angle.Sec[i]]))

    for i in range(len(check)):
        for j in range(i):
            if(check[i]>check[j]):
                b = check[i]
                check[i] = check[j]
                check[j] = b
    second = abs(corr[2])
    a = int((second-math.trunc(second))*n)
    corr1 = [corr[0], corr[1], math.trunc(corr[2])]
    corr2 = [corr[0], corr[1], round_up(corr[2])]
    nth_great = check[a-1] if a != 0 else 0
    if (corr[2]-math.trunc(corr[2]) == 0):
        for i in range (0,n):
            ang = [hz_angle.Deg[i], hz_angle.Min[i], hz_angle.Sec[i]]
            summed = dmsAddition(ang,corr) if corr[0]>=0 and corr[1]>=0 and corr[2]>=0 else dmsSubtraction(ang,ch_sn(corr))
            Deg.append(summed[0])
            Min.append(summed[1])
            Sec.append(summed[2])
            corDeg.append(corr[0])
            corMin.append(corr[1])
            corSec.append(corr[2])
        corr_hz_angle = AngleDMS(Deg, Min, Sec)
    else:
        print(corr1,corr2)
        for i in range(0,n):
            ang = [hz_angle.Deg[i], hz_angle.Min[i], hz_angle.Sec[i]]
            if(Degree[i]>=nth_great):
                summed = dmsAddition(ang,corr2) if corr2[0]>=0 and corr2[1]>=0 and corr2[2]>=0 else dmsSubtraction(ang,ch_sn(corr2))
                corDeg.append(corr2[0])
                corMin.append(corr2[1])
                corSec.append(corr2[2])
            else:
                summed = dmsAddition(ang,corr1) if corr1[0]>=0 and corr1[1]>=0 and corr1[2]>=0 else dmsSubtraction(ang,ch_sn(corr1))
                corDeg.append(corr1[0])
                corMin.append(corr1[1])
                corSec.append(corr1[2])
            Deg.append(summed[0])
            Min.append(summed[1])
            Sec.append(summed[2])
        corr_hz_angle = AngleDMS(Deg, Min, Sec)
    corrections = AngleDMS(corDeg, corMin, corSec)
    return corr_hz_angle ,corrections

def ch_sn(a):
    return [-a[0],-a[1],-a[2]]

def dmsAddition_list(hz_angle):
    sum_angle = [0,0,0]
    for i in range (len(hz_angle.Deg)):
        to_sum = [hz_angle.Deg[i], hz_angle.Min[i], hz_angle.Sec[i]]
        sum_angle = dmsAddition(sum_angle,to_sum)
    return sum_angle

def calculate_lat_dep(leg,bear):
    lat = []
    dep = []
    for i in range (n_major):
        bea = convert_to_deg([bear.Deg[i],bear.Min[i],bear.Sec[i]])
        bea_rad = bea*math.pi/180
        lat.append(round(leg[i]*math.cos(bea_rad),5))
        dep.append(round(leg[i]*math.sin(bea_rad),5))
    return lat , dep

def T_correction_bowditch(leg,ld):
    cor_ld = []
    correc = []
    p = sum(leg)
    er = sum(ld)
    for i in range(n_major):
        corr = round(-(leg[i]/p)*er,5)
        correc.append(corr)
        cor_ld.append(ld[i]+corr)
    return cor_ld , correc


perimeter = 0
for i in range (len(leg_lengths)):
    perimeter += leg_lengths[i]

sum_angle = dmsAddition_list(hz_angle)


th_sum = (n_major-2)*180

th_sum_dms = [th_sum,0,0]
# print(th_sum_dms)

a_error = dmsSubtraction(sum_angle,th_sum_dms)
# print(a_error)

a_correction = [-a_error[0], -a_error[1], -a_error[2]]
# print(a_correction)

a_ind_corr = dmsDivision(a_correction,n_major)
# print(a_ind_corr)

corr_hz_angle , corrections = angleCorrection(hz_angle,a_ind_corr,n_major)
# print(corr_hz_angle.Deg, corr_hz_angle.Min, corr_hz_angle.Sec)

corr_sum_angle = dmsAddition_list(corr_hz_angle)

def print_dms_list(angle):
    for i in range(len(angle.Deg)):
        print(f'{angle.Deg[i]}° {angle.Min[i]}\' {angle.Sec[i]}\"\n')

print('Initial Angles:\n')
print_dms_list(hz_angle)
print('SUM : ', sum_angle)
print('Corrected Angles:\n')
print_dms_list(corr_hz_angle)
print('SUM : ', corr_sum_angle)


bearings = AngleDMS([],[],[])
bearings.Deg.append(bearing[0])
bearings.Min.append(bearing[1])
bearings.Sec.append(bearing[2])


for i in range(1,n_major):
    initial_b = [bearings.Deg[i-1],bearings.Min[i-1],bearings.Sec[i-1]]
    initial_a = [corr_hz_angle.Deg[i],corr_hz_angle.Min[i],corr_hz_angle.Sec[i]]
    bea = calculate_bearing(initial_b,initial_a)
    bearings.Deg.append(bea[0])
    bearings.Min.append(bea[1])
    bearings.Sec.append(bea[2])

a=[corr_hz_angle.Deg[0],corr_hz_angle.Min[0],corr_hz_angle.Sec[0]]
b=[bearings.Deg[n_major-1],bearings.Min[n_major-1],bearings.Sec[n_major-1]]
check_bea = calculate_bearing(a,b)

print('Bearings :\n')
print_dms_list(bearings)

lat, dep = calculate_lat_dep(leg_lengths,bearings)
print(lat,dep)
sum_lat = sum(lat)
sum_dep = sum(dep)
print(sum_lat,'  ',sum_dep)

corr_lat , correction_lat = T_correction_bowditch(leg_lengths,lat)
corr_dep , correction_dep = T_correction_bowditch(leg_lengths,dep)

print(corr_lat,corr_dep)
corr_lat_sum = round(sum(corr_lat),5)
corr_dep_sum = round(sum(corr_dep),5)

print(corr_lat_sum,'  ', corr_dep_sum)

eastings = []
northings = []
eastings.append(easting)
northings.append(northing)

for i in range(1,n_major+1):
    if(i==n_major):
        check_easting = eastings[i-1]+corr_dep[i-1]
        check_northing = northings[i-1]+corr_lat[i-1]
    eastings.append(eastings[i-1]+corr_dep[i-1])
    northings.append(northings[i-1]+corr_lat[i-1])

check_east = eastings[n_major-1] + corr_dep[n_major-1]
check_north = northings[n_major-1] + corr_lat[n_major-1]

def qbToWcb(E,N,quad):
    quad = abs(quad)
    if(E>0 and N>0):
        dms = convert_to_dms(quad)
    if(E>0 and N<0):
        dms = convert_to_dms(180-quad)
    if(E<0 and N>0):
        dms = convert_to_dms(360-quad)
    if(E<0 and N<0):
        dms = convert_to_dms(180+quad)
    return dms

al = []
cb = []
for i in range (n_major):
    if (i == n_major-1):
        dE = eastings[1]-eastings[i]
        dN = northings[1]-northings[i]
    else:
        dE = eastings[i+1]-eastings[i]
        dN = northings[i+1]-northings[i]

    alength = math.sqrt(dE**2+dN**2)
    quad = math.atan(dE/dN)
    wcb = qbToWcb(dE,dN,quad*(180/math.pi))
    al.append(alength)
    cb.append(f'{wcb[0]}° {wcb[1]}\' {wcb[2]}\"')



print('Eastings , Northings :\n')
for i in range(n_major):
    print(eastings[i],' , ',northings[i])


print(check_east,' , ',check_north)

def write_to_column(sheet, column, start_row, no_of_data, data):
    for i in range(0,no_of_data):
        sheet.cell(start_row+i,column).value = data[i]

write_to_column(sheet1, 7 , 4,  n_major , corrections.Deg)
write_to_column(sheet1, 8 , 4 , n_major , corrections.Min)
write_to_column(sheet1, 9 , 4 , n_major , corrections.Sec)
write_to_column(sheet1, 10 , 4 , n_major , corr_hz_angle.Deg)
write_to_column(sheet1, 11 , 4 , n_major , corr_hz_angle.Min)
write_to_column(sheet1, 12 , 4 , n_major , corr_hz_angle.Sec)
write_to_column(sheet1, 13 , 4 , n_major , bearings.Deg)
write_to_column(sheet1, 14 , 4 , n_major , bearings.Min)
write_to_column(sheet1, 15 , 4 , n_major , bearings.Sec)
write_to_column(sheet1, 16 , 4 , n_major , dep)
write_to_column(sheet1, 17 , 4 , n_major , lat)
write_to_column(sheet1, 18 , 4 , n_major , correction_dep)
write_to_column(sheet1, 19 , 4 , n_major , correction_lat)
write_to_column(sheet1, 20 , 4 , n_major , corr_dep)
write_to_column(sheet1, 21 , 4 , n_major , corr_lat)
write_to_column(sheet1, 22 , 4 , n_major , eastings)
write_to_column(sheet1, 23 , 4 , n_major , northings)
write_to_column(sheet1, 24 , 4 , n_major , al)
write_to_column(sheet1, 25 , 4 , n_major , cb)

sheet1.cell(4+n_major,1).value = 'Totals'
sheet1.merge_cells(f'A{4+n_major}:B{4+n_major}')
sheet1.cell(4+n_major,3).value = perimeter
sheet1.cell(4+n_major,4).value = sum_angle[0]
sheet1.cell(4+n_major,5).value = sum_angle[1]
sheet1.cell(4+n_major,6).value = sum_angle[2]
sheet1.cell(4+n_major,10).value = corr_sum_angle[0]
sheet1.cell(4+n_major,11).value = corr_sum_angle[1]
sheet1.cell(4+n_major,12).value = corr_sum_angle[2]
sheet1.cell(4+n_major,13).value = check_bea[0]
sheet1.cell(4+n_major,14).value = check_bea[1]
sheet1.cell(4+n_major,15).value = check_bea[2]
sheet1.cell(4+n_major,16).value = sum_dep
sheet1.cell(4+n_major,17).value = sum_lat
sheet1.cell(4+n_major,20).value = corr_dep_sum
sheet1.cell(4+n_major,21).value = corr_lat_sum
sheet1.cell(4+n_major,22).value = check_east
sheet1.cell(4+n_major,23).value = check_north

for i in range(1,sheet1.max_column+1):
    sheet1.cell(4+n_major,i).alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

wb.save(path)


north = northings
east = eastings
y = []
x = []



pygame.init()

running = True

win_size = [1000,650]
Color_screen=[255,255,255]
Color_line=(255,0,0)

scr = pygame.display.set_mode(win_size)
scr.fill(Color_screen)
# pygame.draw.line(scr,Color_line,(60,80),(130,100))
for i in range(len(north)):
    north[i] = -north[i]

scale = round(500/(max(north)-min(north)),4)

print(scale)
for i in range(len(north)):
    north[i] = north[i] * scale
    east[i] = east[i] * scale

print(north,east)

mn_north = min(north)
mn_east = min(east)


for i in range(len(north)):
    y.append(int(north[i] - mn_north + 100))
    x.append(int(east[i] - mn_east + 50))

print(y,x)



for i in range(len(x)-1):
    pygame.time.delay(1000)
    pygame.draw.line(scr,Color_line,(x[i],y[i]),(x[i+1],y[i+1]),2)
    pygame.draw.circle(scr,Color_line,(x[i],y[i]),5,1)
    pygame.draw.circle(scr,Color_line,(x[i],y[i]),10,1)
    font = pygame.font.Font(r'Muli-Regular.ttf', 16)
    text = font.render(f'M{i}', True, Color_line)
    textRect = text.get_rect()
    textRect.center = (20, 20)
    scr.blit(text,[x[i]-30,y[i]-30])
    pygame.display.flip()

    font = pygame.font.Font(r'Muli-Regular.ttf', 32)
    text = font.render(f'Scale : {scale} pixels = 1 metre', True, Color_line)
    textRect = text.get_rect()
    textRect.center = (100, 50)
    scr.blit(text,[100,5])


path = path.split('.')[0]
path += '.png'
print(path)
try:
    pygame.image.save(scr,path)
except Exception as e:
    print('image save failed')
else:
    print('image saved')
    
while running:

    for events in pygame.event.get():
        if events.type == QUIT:
            running = False
            pygame.quit()

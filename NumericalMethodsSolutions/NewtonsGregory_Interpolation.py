def fact(n):
    x=1
    for i in range(1,n+1):
        x = x*i
    return x
# n =  int(input('\nEnter the value of n : '))
# print('\nEnter the values of x : ')
# for i in range(n):
#     ax.append(float(input()))
# print('\nEnter corresponding values of y : ')
# for i in range(n):
#     ay.append(float(input()))
# x = float(input('\nEnter the value of x for which y is required : '))
n = 4
ax = [10, 20, 30, 40]
ay = [9, 39, 74, 116]
x = 35
diff = []
for i in range(n):
    diff.append([])
for i in range(n):
    diff[i].append(ay[i])
for i in range(1,n):
    for j in range(i,n):
        diff[j].append(round(diff[j][i-1]-diff[j-1][i-1],5))
print('\nDifference Table\n')
for i in diff:
    print(i)
def B_interpolation(ax,ay,diff,n,x):
    h = ax[1]-ax[0]
    xn = ax[-1]
    print(f'xn = {xn}')
    p = (x-xn)/h
    print(f'p = {p}')
    y = diff[-1][0]
    for i in range(1,n):
        py = 1
        for j in range(i):
            py *= (p+j)
        y = y + (py / fact(i)) * diff[-1][i]
    return y
def F_interpolation(ax,ay,diff,n,x):
    h = ax[1]-ax[0]
    x0 = ax[0]
    print(f'x0 = {x0}')
    p = (x-x0)/h
    print(f'p = {p}')
    y = diff[0][0]
    for i in range(1,n):
        py = 1
        for j in range(i):
            py *= (p-j)
        y = y + (py / fact(i)) * diff[i][i]
    return y
y = F_interpolation(ax,ay,diff,n,x) if x < (ax[0]+ax[-1])/2 else B_interpolation(ax,ay,diff,n,x)
print(f'\nThe Value of y at x = {x} is {y}')
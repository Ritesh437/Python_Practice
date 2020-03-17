r='f'

def calc_function(x):
    return x*x*x+2*x*x+4*x-5

print('Interval Start :    ')
a = float(input())
print('Interval End :    ')
b = float(input())
print('Allowed Error :    ')
alllowed_err = float(input())
print('Allowed number of iterations :    ')
max_ite = int(input())
if(calc_function(a)*calc_function(b)<0):
    for i in range(1,max_ite+1):
        mid = (a+b)/2
        if(calc_function(mid)*calc_function(a) < 0):
            b = mid
        else:
            a = mid
        err = abs(b-a)
        print(f'Iteration {i} :    x = {mid}   error = {calc_function(mid)}')
        if(err <= alllowed_err):
            r='t'
            print(f'root = {mid}')
            break
    if(r!='t'):
        print('insufficient iterations')
else:
    print('Invalid Interval')
    

    


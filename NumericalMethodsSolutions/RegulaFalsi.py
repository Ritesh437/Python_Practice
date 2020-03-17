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
        mid = (a*calc_function(b) - b*calc_function(a))/(calc_function(b) - calc_function(a))
        if(calc_function(mid)*calc_function(a) < 0):
            b = mid
        else:
            a = mid
        err = abs(b-a)
        print(f'Iteration {i} :    x = {mid}   error = {err}')
        if(err <= alllowed_err):
            r='t'
            print(f'root = {mid}')
            break
    if(r!='t'):
        print('insufficient iterations')
else:
    print('Invalid Interval')
def calc_function(x):
    return x*x*x*x+4*x*x-6

def calc_diff_function(x):
    return 4*x*x*x+8*x
    
print('Initial value :    ')
a = float(input())
print('Allowed Error :    ')
alllowed_err = float(input())
i=1
y=a
while True:
    x=y
    y = x - (calc_function(x)/calc_diff_function(x))
    print(f'Iteration {i} :    x = {y}    error = {calc_function(y)}\n')
    err = y-x
    i += 1
    if(abs(err) <= alllowed_err):
        print(f'Required Root : {y}')
        break

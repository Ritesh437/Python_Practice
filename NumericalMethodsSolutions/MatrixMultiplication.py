

A = [
    [2,3],
    [3,8],
    [4,5]
]
B = [
    [4,5,2,5],
    [4,8,7,8]
]
print(f'A = {A}')
print(f'B = {B}')

r1 = len(A)
c1 = len(A[0])
r2 = len(B)
c2 = len(B[0])
print(r1,c1,r2,c2)
for i in range(r1):
    row = []
    for j in range(c2):
        sm = 0
        for k in range(r2):
            sm += A[i][k]*B[k][j]
        row.append(sm)
    print(row)
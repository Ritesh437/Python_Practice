import random
import math
rows = 16
cols = 16
mines = 40
board_mine = [['' for _ in range(cols)] for _ in range(rows)]

number = random.sample(range(1,256),40)
for i in range(mines):
    j = number[i]
    x = int(math.trunc(j/16))
    y = int(((j/16)-x)*16)
    board_mine[x][y] = 'M'

for i in board_mine: print(i)
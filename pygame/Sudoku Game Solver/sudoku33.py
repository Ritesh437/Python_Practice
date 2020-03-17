from itertools import *
from copy import copy
import pygame
from pygame.locals import *


def is_distinct(list):
    used = []
    for i in list:
        if i == 0:
            continue
        if i in used:
            return False
        used.append(i)
    return True


def is_valid(board):
    for i in range(3):
        row = []
        col = []
        for j in range(3):
            row.append(board[i][j])
            col.append(board[j][i])
        if not is_distinct(row):
            return False
        if not is_distinct(col):
            return False
    return True


def solve(board,empty):

    if empty == 0:
        return is_valid(board)
    for i in range(3):
        for j in range(3):
            cell = board[i][j]
            if cell != 0:
                continue
            board_new = copy(board)
            for test in [1,2,3]:
                board_new[i][j] = test
                for row in board_new:
                    print(row)
                print('\n')
                if is_valid(board_new) and solve(board_new,empty-1):
                    return True
                board_new[i][j] = 0
    return False


board = [
    [2,0,0],
    [3,1,0],
    [0,2,0]
]

for i in board:
    print(i)
print('\n')

solve(board,5)

for i in board:
    print(i)
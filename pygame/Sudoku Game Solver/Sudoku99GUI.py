from itertools import *
from copy import copy
import pygame
from pygame.locals import *

board = [
    [1, 0, 6,   0, 0, 2,   3, 0, 0],
    [0, 5, 0,   0, 0, 6,   0, 9, 1],
    [0, 0, 9,   5, 0, 1,   4, 6, 2],

    [0, 3, 7,   9, 0, 5,   0, 0, 0],
    [5, 8, 1,   0, 2, 7,   9, 0, 0],
    [0, 0, 0,   4, 0, 8,   1, 5, 7],

    [0, 0, 0,   2, 6, 0,   5, 4, 0],
    [0, 0, 4,   1, 5, 0,   6, 0, 9],
    [9, 0, 0,   8, 7, 4,   2, 1, 0]
]

for i in board:
    print(i)
print('\n')

pygame.init()

running = True

win_size = [550,550]
color_line = [0,0,0]
color_screen = [255,255,255]
font_16 = pygame.font.Font('Muli-Regular.ttf', 16)
font_45 = pygame.font.Font('Muli-Regular.ttf', 40)


scr = pygame.display.set_mode(win_size)
scr.fill(color_screen)
pygame.display.flip()

pygame.draw.rect(scr, color_line,[[50,50], [450,450]], 4)
pygame.display.flip()
width = 2
for i in range(1,9):
    if(i % 3 == 0):
        width = 5
    pygame.draw.line(scr, color_line, [50,50+i*50], [50+450,50+i*50],width)
    pygame.draw.line(scr, color_line, [50+i*50,50], [50+i*50,50+450],width)
    width = 2
    pygame.display.flip()

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
    for i in range(9):
        row = []
        col = []
        for j in range(9):
            row.append(board[i][j])
            col.append(board[j][i])
        if not is_distinct(row):
            return False
        if not is_distinct(col):
            return False
    for i in range(3):
        for j in range(3):
            grid = []
            for k in range(3):
                for l in range(3):
                    grid.append(board[i*3+k][j*3+l])
            if not is_distinct(grid):
                return False
    return True

def find_empty(bo):
    for i in range(len(bo)):
        for j in range(len(bo[0])):
            if bo[i][j] == 0:
                return (i, j)  # row, col
    return None

def solve(board):
    empty = find_empty(board)
    if (not empty):
        return is_valid(board)
    else:
        i, j = empty
    for test in [1,2,3,4,5,6,7,8,9]:
        board[i][j] = test
        for row in board:
            print(row)
        print('\n')
        show_board(board,(255, 69, 56))
        if is_valid(board) and solve(board):
            return True
        board[i][j] = 0
        show_board(board,(255, 69, 56))
    return False


def show_board(board,colour):

    for i in range(9):
        for j in range(9):
            number = board[i][j]
            if(number == 0):
                number = ''
            pygame.draw.rect(scr, color_screen, [[53+j*50,53+i*50],[45,45]])
            pygame.display.flip()
            text = font_45.render(f'{number}', True, colour)
            textRect = text.get_rect()
            textRect.center = (50,50)
            scr.blit(text,[63+j*50,50+i*50])
            pygame.display.update([[54+j*50,54+i*50],[46,46]])


show_board(board,color_line)
pygame.time.delay(3000)
text = font_16.render(f'Solving  ...', True, color_line)
textRect = text.get_rect()
textRect.center = (50,50)
scr.blit(text,[50,15])
pygame.display.flip()

finished = solve(board)

# show_board(board,(255, 69, 56))

if(finished):
    pygame.draw.rect(scr, color_screen, [[50,15],[100,20]])
    text = font_16.render(f'Solved', True, color_line)
    textRect = text.get_rect()
    textRect.center = (50,50)
    scr.blit(text,[50,15])
    pygame.display.flip()
else:
    pygame.draw.rect(scr, color_screen, [[50,15],[100,20]])
    text = font_16.render(f'Unsolvable', True, color_line)
    textRect = text.get_rect()
    textRect.center = (50,10)
    scr.blit(text,[50,15])
    pygame.display.flip()


for i in board:
    print(i)

while running:
    for events in pygame.event.get():
        if (events.type == QUIT):
            running = False
            pygame.quit()



from itertools import *
from copy import copy
import pygame
from pygame.locals import *


board = [
    [2,0,0],
    [3,1,0],
    [0,2,0]
]

for i in board:
    print(i)
print('\n')

pygame.init()

running = True

win_size = [250,250]
color_line = [0,0,0]
color_screen = [255,255,255]
font_16 = pygame.font.Font('Muli-Regular.ttf', 16)
font_45 = pygame.font.Font('Muli-Regular.ttf', 45)

scr = pygame.display.set_mode(win_size)
scr.fill(color_screen)
pygame.display.flip()

pygame.draw.rect(scr, color_line,[[50,50], [150,150]], 2)
pygame.display.flip()
width = 2
for i in range(1,3):
    pygame.draw.line(scr, color_line, [50,50+i*50], [50+150,50+i*50],width)
    pygame.draw.line(scr, color_line, [50+i*50,50], [50+i*50,50+150],width)
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
                show_board(board_new)
                for row in board_new:
                    print(row)
                print('\n')
                if is_valid(board_new) and solve(board_new,empty-1):
                    return True
                board_new[i][j] = 0
                show_board(board_new)
    return False

def show_board(board):

    for i in range(3):
        for j in range(3):
            pygame.draw.rect(scr, color_screen, [[52+j*50,52+i*50],[45,45]])
            pygame.display.flip()
            text = font_45.render(f'{board[i][j]}', True, color_line)
            textRect = text.get_rect()
            textRect.center = (50,50)
            scr.blit(text,[54+j*50,48+i*50])
            pygame.display.flip()
            pygame.time.delay(100)



show_board(board)

finished = solve(board,5)

if(finished):
    text = font_16.render(f'Solved', True, color_line)
    textRect = text.get_rect()
    textRect.center = (50,50)
    scr.blit(text,[50,20])
    pygame.display.flip()
else:
    text = font_45.render(f'Unsolvable', True, color_line)
    textRect = text.get_rect()
    textRect.center = (50,10)
    scr.blit(text,[10,50])
    pygame.display.flip()

while running:

    for events in pygame.event.get():
        if (events.type == QUIT):
            running = False
            pygame.quit()



for i in board:
    print(i)
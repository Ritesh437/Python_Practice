import pygame
from pygame.locals import*
import math
import random
cell_size = 25
rows = 16
cols = 16
mines = 40

pygame.init()
running = True
winSize = [rows*cell_size+100,cols*cell_size+100]
BoardSize = [rows*cell_size,cols*cell_size]
color_scr = [255,255,255]
color_line = [0,0,0]
flag = pygame.image.load(r'flagged.png')
mine = pygame.image.load(r'bomb.png')
blank = pygame.image.load(r'facingDown.png')
boom = pygame.image.load(r'boom.png')

scr = pygame.display.set_mode(winSize)
scr.fill(color_scr)
pygame.display.update()
def set_mine():
    board_mine = [['' for _ in range(cols)] for _ in range(rows)]
    number = random.sample(range(1,256),40)
    for i in range(mines):
        j = number[i]
        x = math.trunc(j/16)
        y = int(((j/16)-x)*16)
        board_mine[x][y] = 'M'
    return board_mine

board = [['' for _ in range(cols)] for _ in range(rows)]
board_mine = set_mine()
for row in board_mine:print(row)


def find_number(j,i):
    mine_count = 0
    for row in range(i-1,i+2):
        for col in range(j-1,j+2):
            if not (row < 0 or row > rows-1 or col < 0 or col > cols-1):
                if(board_mine[row][col] == 'M'): mine_count += 1
    return mine_count

def show_brd():
    for i in range(rows):
        for j in range(cols):
            pygame.draw.rect(scr, color_scr, [[i*cell_size+50,j*cell_size+50], [28,28]])
            if(board[i][j] == 'F'):
                scr.blit(flag, [i*cell_size+50,j*cell_size+50])
            else:
                if(board[i][j] == ''): 
                    scr.blit(blank, [i*cell_size+50,j*cell_size+50])
                else:
                    scr.blit(pygame.image.load(f'{board[i][j]}.png'), [i*cell_size+50,j*cell_size+50])
    pygame.display.flip()

def game_over(i,j):
    scr.blit(mine, [i*cell_size+50,j*cell_size+50])
    for i in range(rows):
        for j in range(cols):
            if board_mine[j][i] =='M':
                scr.blit(mine, [(i)*cell_size+50,(j)*cell_size+50])
                pygame.time.delay(200)
                scr.blit(boom, [(i)*cell_size+50,(j)*cell_size+50])
                pygame.display.update()
    running = False

def start():
    startcell = random.randint(1,80)
    start = 1
    for i in range(rows):
        for j in range(cols):
            if(find_number(i,j) == 0):
                start += 1
            if(start == startcell):
                board[i][j] = 0
                show_brd()
                return

def clicked(x, y, btn):
    if(btn == 1):
        if not board[x-1][y-1] == 'F':
            if(board_mine[y-1][x-1] == 'M'):
                game_over(x-1,y-1)
                return
            else:
                if board[x-1][y-1] == '':
                    board[x-1][y-1] = find_number(x-1, y-1)
    if(btn == 3):
        if(board[x-1][y-1] == 'F'):
            board[x-1][y-1] = ''
        else:
            if(board[x-1][y-1] == ''):
                board[x-1][y-1] = 'F'
            else:
                pass
    show_brd()

start()

while running:
    for events in pygame.event.get():
        if events.type == QUIT:
            running = False
            pygame.quit()
        if events.type == MOUSEBUTTONDOWN:
            pos = pygame.mouse.get_pos()
            x = math.ceil((pos[0] - 50)/cell_size)
            y = math.ceil((pos[1] - 50)/cell_size)
            btn = events.button
            if(x < 1 or x > 16 or y < 1 or y > 16 or btn == 2):
                continue
            clicked(x, y, btn)




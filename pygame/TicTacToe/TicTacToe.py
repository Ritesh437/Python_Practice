import pygame
from pygame.locals import *
import math

b = [
    ['','',''],
    ['','',''],
    ['','','']
]

wins = {
    'X': 0,
    'O': 0
}

turn = 'O'

game = 1

def show_b(b):
    for i in range(3):
        for j in range(3):
            pygame.draw.rect(scr, color_screen, [[52+j*50,52+i*50],[45,45]])
            pygame.display.flip()
            text = font_45.render(f'{b[i][j]}', True, color_line)
            textRect = text.get_rect()
            textRect.center = (50,50)
            scr.blit(text,[59+j*50,51+i*50])
            pygame.display.flip()

pygame.init()

running = True

win_size = [250,250]
color_line = [0,0,0]
color_screen = [255,255,255]
font_16 = pygame.font.Font(r'C:\Windows\Fonts\times.ttf', 16)
font_45 = pygame.font.Font(r'C:\Windows\Fonts\times.ttf', 45)

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

def check_win(b):
    if((b[0][0]=='O' and b[0][1]=='O' and b[0][2]=='O') or (b[0][0]=='X' and b[0][1]=='X' and b[0][2]=='X') or (b[1][0]=='O' and b[1][1]=='O' and b[1][2]=='O') or (b[1][0]=='X' and b[1][1]=='X' and b[1][2]=='X') or (b[2][0]=='O' and b[2][1]=='O' and b[2][2]=='O') or (b[2][0]=='X' and b[2][1]=='X' and b[2][2]=='X') or (b[0][0]=='O' and b[1][0]=='O' and b[2][0]=='O') or (b[0][0]=='X' and b[1][0]=='X' and b[2][0]=='X') or (b[0][1]=='O' and b[1][1]=='O' and b[2][1]=='O') or (b[0][1]=='X' and b[1][1]=='X' and b[2][1]=='X') or (b[0][2]=='O' and b[1][2]=='O' and b[2][2]=='O') or (b[0][2]=='X' and b[1][2]=='X' and b[2][2]=='X') or (b[0][0]=='O' and b[1][1]=='O' and b[2][2]=='O') or (b[0][0]=='X' and b[1][1]=='X' and b[2][2]=='X') or (b[0][2]=='O' and b[1][1]=='O' and b[2][0]=='O') or (b[0][2]=='X' and b[1][1]=='X' and b[2][0]=='X')):
        return True
    else:
        return False

def check_complete(b):
    for i in range(3):
        for j in range(3):
            if(b[i][j]==''):
                return False
    return True

def show_wins():
    pygame.draw.rect(scr, color_screen, [[5,215],[250,45]])
    text = font_16.render(f"Wins =>  X : {wins['X']}     O : {wins['O']}", True, color_line)
    textRect = text.get_rect()
    textRect.center = (50,50)
    scr.blit(text,[5,215])
    pygame.display.flip()

def new_board():
    b = [
        ['','',''],
        ['','',''],
        ['','','']
    ]
    show_b(b)
    turn = 'O' if game % 2 != 0 else 'X'
    pygame.draw.rect(scr, color_screen, [[5,5],[110,45]])
    text = font_16.render(f'Turn : {turn}', True, color_line)
    textRect = text.get_rect()
    textRect.center = (50,50)
    scr.blit(text,[5,5])
    pygame.display.flip()
    show_wins()
    return b,turn

b,turn = new_board()

while running:
    for events in pygame.event.get():
        if (events.type == QUIT):
            running = False
            pygame.quit()
        if (events.type == pygame.MOUSEBUTTONDOWN):
            pos = pygame.mouse.get_pos()
            if(pos[0]>200 or pos[0]<50 or pos[1]>200 or pos[1]<50):
                continue
            i = math.ceil((pos[0]-50)/50) - 1
            j = math.ceil((pos[1]-50)/50) - 1

            if(not b[j][i]==''):
                continue

            b[j][i] = turn
            turn = 'X' if turn =='O' else 'O'
            pygame.draw.rect(scr, color_screen, [[48,5],[47,45]])
            text = font_16.render(f'Turn : {turn}', True, color_line)
            textRect = text.get_rect()
            textRect.center = (50,50)
            scr.blit(text,[5,5])
            pygame.display.flip()

            show_b(b)

            won = check_win(b)
            if(won):
                winner = 'O' if turn == 'X' else 'X'
                pygame.draw.rect(scr, color_screen, [[5,5],[70,45]])
                text = font_16.render(f'Player {winner} Won', True, color_line)
                textRect = text.get_rect()
                textRect.center = (50,50)
                scr.blit(text,[5,5])
                pygame.display.flip()
                game += 1
                if(winner == 'O'):
                    wins['O'] += 1
                if(winner == 'X'):
                    wins['X'] += 1
                show_wins()
                pygame.time.delay(3000)
                b,turn = new_board()

            complete = check_complete(b)
            if(complete):
                pygame.draw.rect(scr, color_screen, [[5,5],[70,45]])
                text = font_16.render(f'Draw', True, color_line)
                textRect = text.get_rect()
                textRect.center = (50,50)
                scr.blit(text,[5,5])
                pygame.display.flip()
                pygame.time.delay(3000)
                game += 1
                b,turn = new_board()

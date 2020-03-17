

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
        if is_valid(board) and solve(board):
            return True
        board[i][j] = 0
    return False


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

solve(board)

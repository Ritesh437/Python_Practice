from cx_Freeze import setup, Executable

base = None    

executables = [Executable("Numbers.py", base=base)]

packages = ["idna"]
options = {
    'build_exe': {    
        'packages':packages,
    },    
}

setup(
    name = "FunWithNumbers",
    options = options,
    version = "1.0.0",
    description = 'Create a list of different type of numbers, save in text file and open a text file',
    executables = executables
)
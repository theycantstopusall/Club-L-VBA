import os

def starter(name, path):
    for root, dirs, files in os.walk(path):
        if name in files:
            return os.path.join(root, name)
    os.open(name)

starter("personalmacro.xlsm", "C://")

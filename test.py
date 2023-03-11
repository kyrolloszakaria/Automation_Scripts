import os

def remove():
    try:
        file_path = 'C:\\Users\\kero6\Desktop\\Noo.txt'
        os.remove(file_path)
    except IOError:
        print('File does not exist')
    print('hello')

print(ascii('k'))
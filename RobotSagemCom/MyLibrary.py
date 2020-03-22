import os
import  sys

class MyLibrary():
    def __init__(self):
        pass

    def send_message(self, message):
        print(message)
    
    def open_file(self , fileName):
        f = open(fileName , 'w')
        f.write('first test \n')
        f.write('second line \n')
        f.close()

if __name__ == "__main__":
    test = MyLibrary()
    test.send_message("This is a message")
    test.open_file('test.txt')
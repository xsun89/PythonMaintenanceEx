import os
import sys


while True:
    name = raw_input("your name")
    if name == 'xsun':
        psw = '123'
        inputPsw = raw_input("your password")
        while inputPsw != psw:
            inputPsw = raw_input("your password")
        break;
    else:
        print "wrong name"

while True:
    searchName = raw_input("Please input the name you want to find")
    searchfile = open("contact.txt", "r")
    flag = False
    for line in searchfile:
        if searchName in line:
            print line
            flag = True
            continue
    if flag == False:
        print 'no match found'

    searchfile.close()



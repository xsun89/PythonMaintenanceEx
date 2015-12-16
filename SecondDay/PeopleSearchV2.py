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

searchfile = open("contact.txt", "r")
contactDict = {}
for line in searchfile:
    lineSplit = line.split();
    contactDict[lineSplit[1]] = line

searchfile.close()
print contactDict.items()
while True:
    searchName = raw_input("Please input the name you want to find")
    if contactDict.has_key(searchName):
        print contactDict.get(searchName)
    else:
        print 'no match found'




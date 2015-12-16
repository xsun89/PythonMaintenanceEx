import os
product = ["abc", "cdf", "efg"]
print product.index("cdf")
f = open('contact.txt')
while True:
    line = f.readline();
    afterSplit = line.split();
    print afterSplit
    if(len(line) == 0):
        break
f.close()

import fileinput
for line in fileinput.input('contact.txt'):
    line =line.replace('DaiQingYang', 'DaiQingYang2222222')
    print line


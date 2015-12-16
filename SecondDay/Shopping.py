productItem = []
productPrice = []
shopItem = []

f = file('shopItem.txt')
while True:
    line = f.readline();
    if len(line) == 0:
        break

    shopItemList = line.split()
    productItem.append(shopItemList[0])
    productPrice.append(int(shopItemList[1]))

print productItem
print productPrice

salary = int(raw_input("Please enter your salary"))
while True:
    cnt = 0
    for p in productItem:
        idx = p.index(p);
        print cnt, ':', p, '\t', productPrice[idx]
        cnt = cnt + 1
    print cnt, ':exit'
    ret = raw_input("Please select:")
    if int(ret) > cnt:
        break

    shopItem.append(int(ret))
    print shopItem

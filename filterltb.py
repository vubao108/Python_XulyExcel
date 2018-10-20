mydict = {}
with open('loaitb') as f:
    for line in f:
        line = line.rstrip()
        mydict[line] = mydict.get(line,0) + 1

for key in mydict:
    print key
import re
f = open('dd1.txt').readlines()
f1 = open('dd.txt','w')
for i in f:
    arr = i.split('((')
    arr1 = arr[1].split('))')
    y = arr1[0] +'\n'
    print (y)
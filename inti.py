# coding=utf-8
import makeURL
import os

path = os.path.abspath("init.py")
divide = path.split("/")

newPath = ""
for i in range(0, len(divide)-1):
    newPath = newPath + "/" + divide[i]

pathList = os.listdir(newPath)

j = 0
for line in pathList:
    if "interface" in line:
        j += 1
        makeURL.start(line, j)

if j == 0:
    print "No interface file."

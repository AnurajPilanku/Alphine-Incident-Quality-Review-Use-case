#Anuraj Pilanku
#Alphine usecase
import os
import sys

path =sys.argv[1]# "D:/Pycharm projects/GeeksforGeeks/Nikhil"
dir = os.listdir(path)
if len(dir) == 0:
	print("Empty directory")
else:
	print("Not empty directory")

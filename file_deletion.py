#Anuraj Pilanku

import os
import sys
dir=sys.argv[1]

for f in os.listdir(dir):
    os.remove(os.path.join(dir,f))

file_paths = open(sys.argv[2], 'r').read().split("\n")
for file_path in file_paths:
    if(os.path.exists(file_path)):
        os.remove(file_path)
print("success")


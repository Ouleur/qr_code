import pyqrcode 
import png 
from pyqrcode import QRCode 
import pandas as pd
import os
import sys
from openpyxl import *

total = len(sys.argv)
cmdargs = str(sys.argv)

print ("Args list: %s " % cmdargs)


def create_directory(path):
    # Check whether the specified path exists or not
    isExist = os.path.exists(path)
    if not isExist:
        # Create a new directory because it does not exist 
        os.makedirs(path)
        print("The new directory is created!")

create_directory(sys.argv[1])
data = pd.read_excel (sys.argv[2])
df = pd.DataFrame(data, columns= ['MATRICULE'])
df = df.values.tolist()


wb = load_workbook(filename=sys.argv[2])
ws = wb.worksheets[0]


i =2
for x in df:
    print (x)
    url = pyqrcode.create(x[0]) 
    url.png('{}/{}.png'.format(sys.argv[1],x[0]), scale = 6) 
    # row number = 0 , column number = 1
    ws.cell(row=i, column=int(sys.argv[4])).value = '{}.png'.format(x[0])
    i+=1
    
# save the file
wb.save(sys.argv[3])

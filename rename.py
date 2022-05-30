
import csv
import shutil

from os import path

import os, glob, PIL, qrcode, openpyxl
from PIL import Image, ImageDraw, ImageFont, PngImagePlugin


PIL.PngImagePlugin.MAX_TEXT_CHUNK= 10485760
PIL.PngImagePlugin.MAX_TEXT_MEMORY= 971088604

Image.MAX_IMAGE_PIXELS = None

path_folder= input("input path here: ")

list_png = glob.glob(path_folder+'/**/*.png', recursive=True)
list_jpg = glob.glob(path_folder+'/**/*.jpg', recursive=True)
list_csv = glob.glob(path_folder+'/**/*.csv', recursive=True)



list_ds =[]

# tao folder file final
path_final = path_folder + "/final"
if os.path.exists(path_final) == False: 
    os.mkdir(path_final)




# lay list design
for p in list_png:
    list_ds.append(p)

for j in list_jpg:
    list_ds.append(j)

print(list_csv[0])

with open(list_csv[0]) as f:
    # print(f.read())
    reader = csv.reader(f)
    i  = [row for row in reader]
    
    i_columns =[list(x) for x in zip(*i)][1]

    # print(i_columns)



for ds in list_ds:
    #detect info product
    #img = Image.open(ds)
    ds_name= os.path.splitext(os.path.basename(ds))[0]
    ds_ex=os.path.splitext(os.path.basename(ds))[1]

    
    row_num = i_columns.index(ds_name)
    
    block_newname = str(i[row_num][0])
    shutil.copyfile(ds, path_final+"/"+ block_newname + ds_ex)

    #img.save(path_final+"/"+block_newname+ds_ex)
    print(block_newname + "   done")

print("finish..........")


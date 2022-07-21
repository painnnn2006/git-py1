
import csv
import shutil

from os import path

import os, glob, PIL, qrcode, openpyxl
from PIL import Image, ImageDraw, ImageFont, PngImagePlugin

import shutil, zipfile
from zipfile import ZipFile
from shutil import make_archive


PIL.PngImagePlugin.MAX_TEXT_CHUNK= 10485760
PIL.PngImagePlugin.MAX_TEXT_MEMORY= 971088604

Image.MAX_IMAGE_PIXELS = None

path_folder= input("input path here: ")
code_run = input("rename or filter: ")
list_png = glob.glob(path_folder+'/**/*.png', recursive=True)
list_jpg = glob.glob(path_folder+'/**/*.jpg', recursive=True)
list_csv = glob.glob(path_folder+'/**/*.csv', recursive=True)
list_xlsx = glob.glob(path_folder+'/**/*.xlsx', recursive=True)
list_fin = []


list_ds =[]






# lay list design
for p in list_png:
    list_ds.append(p)

for j in list_jpg:
    list_ds.append(j)

#print(list_csv[0])
if code_run == 1:
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

    
    if code_run == "1":
        # tao folder file final
        path_final = path_folder + "/final"
        if os.path.exists(path_final) == False: 
            os.mkdir(path_final)
        row_num = i_columns.index(ds_name)
    
        block_newname = str(i[row_num][0])
        shutil.copyfile(ds, path_final+"/"+ block_newname + ds_ex)

        #img.save(path_final+"/"+block_newname+ds_ex)
        print(block_newname + "   done")
    elif code_run == "2":

        ds_final_name  = ds_name.split('_')[0] + '-' + ds_name.split('_')[2].split('.')[0]
        
        folder_name = ds_name.split('.')[1]
        path_final = path_folder + "/" + folder_name
        if os.path.exists(path_final) == False: 
            os.mkdir(path_final)
            list_fin.append(path_final)
        shutil.copyfile(ds, path_final+"/"+ ds_final_name + ds_ex)

        img = Image.open(path_final+"/"+ ds_final_name + ds_ex)
        d = ImageDraw.Draw(img)
        fnt= ImageFont.truetype("arial.ttf",80)
        d.text((2900,4900),ds_final_name,font= fnt, fill=(0,0,0))
        img.save(path_final+"/"+ ds_final_name + ds_ex, dpi=(300,300))
        print(ds_final_name + "done!!!")
  #move xlsx
for xlsx in list_xlsx:
            xlsx_name = os.path.splitext(os.path.basename(xlsx))[0]
            path_xlsx_final = path_folder + "/" + xlsx_name
            shutil.move(xlsx,path_xlsx_final)
for fin in list_fin:
            shutil.make_archive(fin,"zip", fin)
            print(fin.split('/')[1], "done")








            
            
    

print("Finish..........")


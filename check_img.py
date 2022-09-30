#from PIL import Image
#import operator

#path_img = input("path img: " )
#img = Image.open(path_img).convert('1')
#black,white = img.getcolors()
#print(white[0])
#print(black[0])



# 2


import numpy
import os, glob, PIL, qrcode, openpyxl

from os import path
from PIL import Image, ImageDraw, ImageFont, PngImagePlugin
from distutils.log import debug
import logging




PIL.PngImagePlugin.MAX_TEXT_CHUNK= 1048576
PIL.PngImagePlugin.MAX_TEXT_MEMORY= 97108864

Image.MAX_IMAGE_PIXELS = None



def get_image(image_path):
    """Get a numpy array of an image so that one can access values[x][y]."""
    
    list_png = glob.glob(path_folder+'/**/*.png', recursive=True)
    list_jpg = glob.glob(path_folder+'/**/*.jpg', recursive=True)
    list_ds =[]

    for p in list_png:
        list_ds.append(p)

    for j in list_jpg:
        list_ds.append(j)
    #create new xlsx
    wb = openpyxl.Workbook()
    ws = wb.active
    i = 0

    #check each ds
    for ds in list_ds:
        image = Image.open(ds, "r")
        ds_name = os.path.splitext(os.path.basename(ds))[0]
        width, height = image.size
        pixel_values = list(image.getdata())
        tran_values = []
        for px in pixel_values:
            #print(px)
            if image.mode == "RGBA" and  px[3] == 0  :
                tran_values.append(px)
                #print(px)
               
            elif image.mode == "RGBA" and px[0] == 255 and px[1] == 255 and px[2] == 255 and px[3] == 1:
                tran_values.append(px)
            elif image.mode == "RGB" and px[0] == 255 and px[1] == 255 and px[2] == 255:
                tran_values.append(px)
                
        rate_trans = len(tran_values) / len(pixel_values)
    
        #fill data in xlsx 
        ws.cell(column=1, row = i+1, value=ds_name)
        ws.call(column=2, row = i+1, value=rate_trans)
        i = i+1
        print(ds_name + "check done")

    ws.save(path_folder + "/check_px_trans.xlsx", recursive=true)

    return "Done " + len(list_ds) + " ds"


path_folder = input("path folder: " )
image = get_image(path_img)
print(image)
#print(len(tran_values))
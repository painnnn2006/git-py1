import numpy as np
import os, glob, qrcode, openpyxl
from PIL import Image, ImageDraw, ImageFont
from os import path
import shutil, zipfile
from zipfile import ZipFile
from shutil import make_archive

#anhvldaide
#hello world



# load database
path_folder= input("input path here: ")
list_png = glob.glob(path_folder+'/**/*.png', recursive=True)
list_jpg = glob.glob(path_folder+'/**/*.jpg', recursive=True)
list_xlsx = glob.glob(path_folder+'/**/*.xlsx', recursive=True)
list_ds =[]
list_fin = []
for p in list_png:
    list_ds.append(p)

for j in list_jpg:
    list_ds.append(j)

list_ds.sort()

for ds in list_ds:
    #detect info product
    ds_name= os.path.splitext(os.path.basename(ds))[0]
    ds_ex = ds_ex=os.path.splitext(os.path.basename(ds))[1]
    ds_sp = ds_name.split('_')[2]
    ds_dpi =int(ds_name.split('.')[3])
    ds_w_s = int(ds_name.split('.')[1])
    ds_h_s=int(ds_name.split('.')[2])
    ds_final_name= ds_name.split('_')[0] + "_" + ds_name.split('_')[1] 
    folder_name = ds_name.split('_')[3].split('.')[0]
    path_final = path_folder + "/" + folder_name
    if os.path.exists(path_final) == False: 
        os.mkdir(path_final)
        list_fin.append(path_final)
    


    # print(ds_final_name, "...", folder_name)

    # print(ds_dpi)
    print(ds_final_name.split('_')[0]," begin render...")
    # print(path_final)
    im=Image.open(ds)
    if ds_sp == 'WC':
        # print(im.mode)
        
        path_RGB = path_final + "/RGB"
        # remove white back ground
        # img = im.convert("RGBA")
        # datas = img.getdata()
        # newData = []
  
        # for items in datas:
        #     if items[0] == 255 and items[1] == 255 and items[2] == 255:
        #         newData.append((255, 255, 255, 0))
        #     else:
        #         newData.append(items)
  
        # img.putdata(newData)
        if im.mode == "RGB":
            new = Image.new("RGB",(3266,1335),(0,0,0))
            new.paste(im,(13,19))
            if os.path.exists(path_RGB) == False: 
                os.mkdir(path_final)

            new.save(path_RGB+"/"+ds_final_name+'.PNG', dpi=(ds_dpi,ds_dpi))


        if im.mode == "RGBA":
            new = Image.new("RGBA",(3266,1335),(0,0,0,0))
            new.paste(im,(13,19))
            new.save(path_final+"/"+ds_final_name+'.PNG', dpi=(ds_dpi,ds_dpi))



    # print(ds_sp)
    # if ds_sp == 'SO':
    #     img_r = im.resize((ds_w_s,ds_h_s)).transpose(Image.FLIP_LEFT_RIGHT).transpose(Image.ROTATE_180)
    # else:
    if ds_sp == 'WAL':
        img_w = im.resize((ds_w_s,ds_h_s))
        img_r = Image.new(img_w.mode,(ds_w_s,ds_h_s + 47),color=(255,255,255))
        d = ImageDraw.Draw(img_r)
        fnt= ImageFont.truetype("arial.ttf",30)
        d.text((ds_w_s/2+1000,ds_h_s),ds_final_name,font= fnt, fill=(0,0,0))
        img_r.paste(img_w,(0,0))        
        img_r.save(path_final+"/"+ds_final_name+"."+ ds_ex, dpi=(ds_dpi,ds_dpi))
    

    #     else:

    #         img_r = im.resize((ds_w_s,ds_h_s)).transpose(Image.FLIP_LEFT_RIGHT)

    # # print(img_r.mode)
    # im_c = img_r.convert('RGB')

    # print(im_c.mode)
# #     # Read input image, and convert to NumPy array. 
#     img = np.array(img_r)  # img is 1080 rows by 1920 cols and 4 color channels, the 4'th channel is alpha.

# #     # Find indices of non-transparent pixels (indices where alpha channel value is above zero).
#     idx = np.where(img[:, :, 3] > 0)

# #     # Get minimum and maximum index in both axes (top left corner and bottom right corner)
#     x0, y0, x1, y1 = idx[1].min(), idx[0].min(), idx[1].max(), idx[0].max()

# #     # Crop rectangle and convert to Image
#     out = Image.fromarray(img[y0:y1+1, x0:x1+1, :])

#     # print(path_final+"/"+ds_final_name+ds_ex)
#     # out.show()

    

    

    print(ds_final_name.split('_')[0]," Done")

# move xlsx + zip 

# xlsx
for xlsx in list_xlsx:
    xlsx_name = os.path.splitext(os.path.basename(xlsx))[0]
    path_xlsx_final = path_folder + "/" + xlsx_name
    # move
    # print(xlsx)
    # print(xlsx_name)
    # print(path_xlsx_final)
    shutil.move(xlsx,path_xlsx_final)

# zip final folde

for fin in list_fin:
    shutil.make_archive(fin,"zip", fin)
    print(fin, "done")




    
print("Finish. Good luck have fun:) ")













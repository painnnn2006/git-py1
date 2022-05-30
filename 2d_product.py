import numpy as np
import os, glob, qrcode, openpyxl
import PIL
from PIL import Image, ImageDraw, ImageFont, PngImagePlugin

from os import path

PIL.PngImagePlugin.MAX_TEXT_CHUNK= 1048576
PIL.PngImagePlugin.MAX_TEXT_MEMORY= 97108864

# load database
path_folder= input("input path '''''2D PRODUCT''''' here: ")
list_ds = glob.glob(path_folder+'/**/*.png', recursive=True)
for ds in list_ds:
    #detect info product
    ds_name= os.path.splitext(os.path.basename(ds))[0]
    ds_ex=os.path.splitext(os.path.basename(ds))[1]
    ds_sz = ds_name.split('-')[0]
    ds_sp = ds_name.split('-')[1].split('_')[4]
    ds_dpi =150
    ds_w_s = int(ds_name.split('.')[1])
    ds_h_s=int(ds_name.split('.')[2])
    ds_final_name= ds_name.split('.')[0]
    folder_name = ds_name.split('.')[3]
    dup_ds = int(ds_name.split('.')[4])


    print(ds_final_name," begin render...")
    im=Image.open(ds)
    img_r = im.resize((ds_w_s,ds_h_s))
    
#     # Read input image, and convert to NumPy array. 
    img = np.array(img_r)  # img is 1080 rows by 1920 cols and 4 color channels, the 4'th channel is alpha.

#     # Find indices of non-transparent pixels (indices where alpha channel value is above zero).
    idx = np.where(img[:, :, 3] > 0)

#     # Get minimum and maximum index in both axes (top left corner and bottom right corner)
    x0, y0, x1, y1 = idx[1].min(), idx[0].min(), idx[1].max(), idx[0].max()

#     # Crop rectangle and convert to Image
    out = Image.fromarray(img[y0:y1+1, x0:x1+1, :])

    # print(path_final+"/"+ds_final_name+ds_ex)
    # out.show()

    path_final = path_folder + "/" + folder_name

    if os.path.exists(path_final) == False: 
        os.mkdir(path_final)

    #save + dup ds
    if dup_ds == 1 :
        out.save(path_final+"/"+ds_final_name+ds_ex, dpi=(ds_dpi,ds_dpi))
    elif dup_ds >1:   
        for a in range(1,dup_ds):

            out.save(path_final+"/"+ds_final_name+'_' + a +ds_ex, dpi=(ds_dpi,ds_dpi))

    print(ds_final_name," Done")

    
print("Finish. Good luck have fun:) ")













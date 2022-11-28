import os, glob
import shutil
from os import path


path_folder= input("input path ''''''DS'''''''' here: ")
list_csv = glob.glob(path_folder+'/**/*.csv', recursive=True)
path_eror_folder = path_folder + "/eror"

# tao folder file error
path_final = path_folder + "/fix"
if os.path.exists(path_final) == False: 
	os.mkdir(path_final)

#get list file error
list_png = glob.glob(path_eror_folder +'/**/*.png', recursive=True)

for ds in list_png:
	ds_name = os.path.splitext(os.path.basename(ds))[0]
	ds_move = path_folder + "/" + ds_name
	#move ds eror to fix
	path_final_ds = path_final + "/" + ds_name
	shutil.move(ds_move,path_final_ds)
	
#shutil.rmtree(path_eror_folder)
print("Done")

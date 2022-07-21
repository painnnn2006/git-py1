from typing import _T_co
import numpy as np
import os, glob, PIL, qrcode, openpyxl

from os import path
import shutil, zipfile
from zipfile import ZipFile
from shutil import make_archive


path_folder= input("input path folder lot here: ")
list_xlsx = glob.glob(path_folder+'/**/*.xlsx', recursive=True)

for lot in list_xlsx:
    wbinput = openpyxl.load_workbook(lot)
    isheet = wbinput.sheetnames
  t_sheet = wbinput.get_sheet_by_name('TOTAL')

    for sheet in isheet:
        if sheet == "SINGLE":
            s_sheet = wbinput.get_sheet_by_name('SINGLE')

            t_sheet["A1"].value = 
                    
        elif sheet == "MULTI":
            m_sheet = wbinput['MULTI']

        elif sheet == "TOTAL":
              




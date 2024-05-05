'''import pandas as pd
from openpyxl import Workbook, worksheet
from  openpyxl.drawing.image import Image
xls = pd.ExcelFile('Project1.xls')
sheetX = xls.parse(0) # 0 is the sheet number
print(sheetX)



import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.drawing.image import Image


openpyxl_version = openpyxl.__version__
print(openpyxl_version)  #to see what version I'm running

# downloaded a .png to local directory manually from
# "https://www.python.org/static/opengraph-icon-200x200.png"

#change to the location and name of your image
png_loc = r'Row 17.png'

# test.xlsx already exists in my current directory 

wb = load_workbook('Project1.xlsx')
ws = wb.active
my_png = openpyxl.drawing.image.Image(png_loc)
ws.add_image(my_png, 'D3')
wb.save('test.xlsx')'''

from PIL import Image
import xlwt
from io import BytesIO
from xlutils.copy import copy
from xlrd import open_workbook
import pandas as pd
file = open_workbook('Project1.xls')
file_sheet = file.sheet_by_index(0)
workbook   = copy(file)
worksheet1 = workbook.get_sheet(0)
worksheet1.row(4).height_mismatch = True
worksheet1.row(4).height = 1000
img = Image.open("Row 17.png")
image_parts = img.split()
r = image_parts[0]
g = image_parts[1]
b = image_parts[2]
img = Image.merge("RGB", (r, g, b))
fo = BytesIO()
img.save(fo, format='bmp')
worksheet1.insert_bitmap_data(fo.getvalue(),4,4)
workbook.save('Project1.xls')
img.close()

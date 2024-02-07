from openpyxl import Workbook
from openpyxl.drawing.image import Image



wb = Workbook()
ws = wb.active
# create an image
img = Image('/media/tosin/7859-B0EF/TemplateImg.png')

INCHES_TO_PIXEL = 0.0104145

img.width = 1.67 / INCHES_TO_PIXEL
img.height = 2.63 / INCHES_TO_PIXEL
ws.add_image(img, 'A1')
# add to worksheet and anchor next to cells
wb.save('logo.xlsx')

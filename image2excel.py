#!/usr/bin/env python3
import xlsxwriter
from PIL import Image
import sys

def componentToHex(component):
    if len(str(hex(component)[2:])) == 1:
      return '0' + hex(component)[2:]
    else:
      return hex(component)[2:]

def rgbToHexa(rval, gval, bval):
    return '#' + componentToHex(rval) + '' + componentToHex(gval) + '' + componentToHex(bval)

def main():

    if len(sys.argv) < 2:
        sys.exit('No input image!')
    # load image
    im = Image.open(sys.argv[1])
    imsize = im.size
    if imsize[0] > 128 or imsize[1] > 128:
        im.thumbnail((128,128), Image.ALIAS) # downsampling, aspect ratio stays the same
        imsize = im.size
    pix = im.load()

    # create excel workbook/worksheet
    workbook = xlsxwriter.Workbook( './output.xls')
    worksheet = workbook.add_worksheet('image')

    worksheet.set_column(0, imsize[0], 3);

    # color excel cells
    for x in range(imsize[0]):
        for y in range(imsize[1]):
            colors = pix[x, y]
            hexcode = rgbToHexa(colors[0], colors[1], colors[2])
            wbformat = workbook.add_format({'bg_color':hexcode})
            worksheet.write(y,x,'', wbformat)

    workbook.close()

if __name__ == '__main__':
    main()

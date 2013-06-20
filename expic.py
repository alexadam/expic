#!/usr/bin/python
from __future__ import division
from PIL import Image
from xlsxwriter.workbook import Workbook
import sys, getopt

def start(inputfile):
    workbook = Workbook(inputfile + '.xlsx')
    worksheet = workbook.add_worksheet()
    
    im = Image.open(inputfile)
    width, height = im.size
    rgb_im = im.convert('RGB')
    
    for x in range(width):
        for y in range(height):
            red, green, blue = rgb_im.getpixel((x, y))
                
            pixel_hex = format((red<<16) | (green<<8) | blue, '06x')
            
            worksheet.write(y, x, '', workbook.add_format({'bg_color' : str(pixel_hex)}))
    
    workbook.close()

if __name__ == '__main__':
    inputfile = ''
     
    try:
        opts, args = getopt.getopt(sys.argv[1:], "hi:r:")
    except getopt.GetoptError:
        print 'expic.py -i <input_picture>'
        sys.exit(2)
         
    for opt, arg in opts:
        if opt == '-h':
            print 'expic.py -i <input_picture>'
            sys.exit()
        elif opt == "-i":
            inputfile = arg

    start(inputfile)

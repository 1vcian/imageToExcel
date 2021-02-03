from PIL import Image
import xlsxwriter

def rgb2hex(r,g,b):
    return "#{:02x}{:02x}{:02x}".format(r,g,b)

while(1):
    fileName=input("input filename with extension: ")
    im = Image.open(fileName).convert('RGBA')
    background = Image.new('RGBA', im.size, (255,255,255))
    rgb_image = Image.alpha_composite(background, im).convert('RGB')
    width, height = im.size
    workbook = xlsxwriter.Workbook(fileName[:fileName.rfind('.')]+'.xlsx')
    worksheet = workbook.add_worksheet()
    for x in range(width):
        for y in range(height):
            r,g,b = rgb_image.getpixel((x,y))
            c=rgb2hex(r,g,b)
            cell_format = workbook.add_format({'text_wrap': True})
            cell_format.set_bg_color(c)
            cell_format.set_font_color('#aaaaaa' if r+g+b<255 else '#000000')
            data=str(c)+' ('+str(x)+','+str(y)+')'
            worksheet.write(y,x,'' if c=='#ffffff' else data,cell_format)
    workbook.close()

import xlrd

from xlrd import open_workbook

book = xlrd.open_workbook("book1.xlsx")
sheet = book.sheet_by_index(0)

original_list = []
clean_list = []

for k in range(1,sheet.nrows):
    original_list.append(str(sheet.row_values(k)))

print(original_list)


#Removing Special Character from the Data using Regex
import re
#cleanString = re.sub('\W+'," ", "['Neel Kamath']" )
for i in original_list:
    clean_string = re.sub('[^A-Za-z0-9]+', " ", i)
    clean_list.insert( 0 , clean_string)
    
print (clean_list) 


from PIL import Image, ImageDraw, ImageFont
  
def hacker_name(name):
    image = Image.open('certificate_raw.jpg')   # create Image object with the input image
    draw = ImageDraw.Draw(image)    # initialise the drawing context with the image object as background
    font = ImageFont.truetype('Roboto-Bold.ttf', size=100)

    if len(j) < 12:
        (x , y) = (2090, 1520)

    elif len(j) < 20:
        (x , y) = (1950, 1520)
            
    else:
        (x , y) = (1850, 1520)
                 
    color = 'rgb(0, 0, 0)' # black color
    draw.text((x, y), name, fill=color, font=font)
    image.save(name + ".jpg")  # save the edited image

    #image.save('optimized.png', optimize=True, quality=20)

for j in clean_list:
    hacker_name(j)
    print(len(j))
    print (j)
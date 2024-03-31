import pyautogui
import time
import subprocess
import os
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import openpyxl
from openpyxl.styles import  Alignment
from openpyxl.drawing.image import Image as XLImage

# Set the path to the Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


#coordinates of calculator button on gui interface (proteus)
buttons_coordinates = {
  '1':(860,473),'2':(918,473),'3':(973,473),'+':(1036,473),
  '4':(860,497),'5':(918,497),'6':(973,497),'-':(1036,497),
  '7':(860,522),'8':(918,522),'9':(973,522),'*':(1036,522),
  '.':(860,550),'0':(918,550),'=':(973,550),'/':(1036,550),
}


# openning the excell sheet that contains the testcases written
wb = openpyxl.load_workbook(r'test_cases.xlsx')
center_aligned_text = Alignment(horizontal="center")
sheet = wb.active

# reading the test cases in a list
testcases_col = []
for cell in sheet['A'][1:]:
  testcases_col.append(cell.value)



# openning protues
subprocess.Popen('c://Program Files (x86)//Labcenter Electronics//Proteus 8 Professional//BIN//PDS.EXE')
time.sleep(3)
#openning the caculator project
pyautogui.click(393,215,duration=1)
time.sleep(1)
pyautogui.write("d:\Embedded_Systems\AVR_Workspace\\NeuronetiX-Advanced-Calculator\\NeuronetiX_Advanced_Calculator.pdsprj")
time.sleep(1)
pyautogui.click(x=994, y=573,duration=1)
time.sleep(5)

#looping on test cases
for i,tc in enumerate(testcases_col):
  #run the simulation
  pyautogui.click(x=24, y=717,duration=1)
  time.sleep(2) 
  # process each test case
  for c in testcases_col[i]:  
    pyautogui.click(buttons_coordinates[c],duration=0.5)
    time.sleep(0.5)

  pyautogui.click(buttons_coordinates['='],duration=0.5)
  time.sleep(5)

  #taking screenshot for the result on proteus
  screenshot = pyautogui.screenshot()
  region = screenshot.crop((895, 275, 1000, 296))  #coordinates
  region.save("temp{}.png".format(i))

  # Open the image
  #image = Image.open("temp{}.png".format(i))
  
  # Enhance image contrast
  #enhancer = ImageEnhance.Contrast(image)
  #image = enhancer.enhance(1.5)  # Increase contrast by a factor of 2
  
  # Convert to grayscale
  #image = image.convert("L")

  # Apply thresholding
  #threshold = 150
  #image = image.point(lambda p: p < threshold and 255)
  #image.save("temp{}.png".format(i))
  
  #Extract text using Tesseract OCR
  #custom_config = r'--oem 1 --psm 6 -c tessedit_char_whitelist=0123456789.-'
  #extracted_text = pytesseract.image_to_string(image, lang='fra' , config=custom_config)
  
  #print(extracted_text)




  # Create a placeholder for the image in the worksheet
  xl_img = XLImage("temp{}.png".format(i))
  xl_img.anchor = 'C{}'.format(i+2)  # Specify the cell where the top-left corner of the image will be placed
  sheet.add_image(xl_img)


  #sheet['C{}'.format(i+2)] = image
  #sheet['C{}'.format(i+2)].alignment = center_aligned_text
  wb.save(r'test_cases.xlsx')

  #stop the simulation
  pyautogui.click(x=133, y=717,duration=0.5)
  time.sleep(1) 
















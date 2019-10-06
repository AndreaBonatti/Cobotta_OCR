import win32com.client
from PIL import Image
import pytesseract
import cv2
import os
# import sys


# Setup ORiN vars
cao_engine = win32com.client.Dispatch("CAO.CaoEngine")
controller = self.cao_engine.Workspaces(0).AddController("N10-W02", "CaoProv.Canon.N10-W02", "", "Conn=eth:192.168.0.90")
picture = self.controller.AddVariable("IMAGE")
# Take picture
self.controller.Execute("OneShotFocus")
# Note: imageByteArray is already an encoded PNG
imageByteArray = self.picture.Value
newFile = open("pic.png", "wb")
newFile.write(imageByteArray)

image = cv2.imread("pic.png", 1)
gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

# Code to use the OCR on a test image
# argc = len(sys.argv)
# if argc > 1:
#     image = cv2.imread(sys.argv[1], 1)
# else:
#     print("Errore! Nessuna immagine inserita!")
#
# gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

# write the grayscale image to disk as a temporary file so we can apply OCR to it
filename = "{}.png".format(os.getpid())
cv2.imwrite(filename, gray)

# load the image as a PIL/Pillow image, apply OCR, and then delete the temporary file
text = pytesseract.image_to_string(Image.open(filename))
os.remove(filename)
# print the text that is read by the OCR
print(text)

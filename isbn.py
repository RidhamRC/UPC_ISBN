import cv2
from pyzbar.pyzbar import decode
from openpyxl import Workbook
import os
from PIL import Image
from pyzbar.wrapper import ZBarSymbol
import pytesseract
import re
def extract_barcodes(image_path):
    # Load the image
    upc = None
    isbn = None
    isbn_matches=None
    
    img = Image.open(image_path)
    reg1=r'\d{1}-?\d{5}-?\d{3}-\s*?[0-9X]{1}'
    reg2=r'\d{1}-?\d{4}-?\d{4}-\s*?[0-9X]{1}'
    pytesseract.pytesseract.tesseract_cmd = "Tesseract-OCR\\tesseract.exe"
    # Use Tesseract to do OCR on the image
    text = pytesseract.image_to_string(img)
    text=text.replace("\n"," ")
    
    isbn_matches = re.findall(reg1, text)
    print(isbn_matches,"FOR REG1",image_path)
    if not isbn_matches:
        isbn_matches = re.findall(reg2, text)
        print(isbn_matches,"FOR REG2",image_path)
    
    if isbn_matches:
        isbn=isbn_matches[0]
    
    
    image = cv2.imread(image_path)
    # Convert the image to grayscale
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    
    # Use Pyzbar to decode barcodes
    barcodes = decode(gray_image,[ZBarSymbol.EAN13])
    
    
    
    # Iterate through detected barcodes
    for barcode in barcodes:
        barcode_data = barcode.data.decode("utf-8")
        barcode_type = barcode.type
        # print(image_path,len(barcode_data),barcode_type)
        upc = barcode_data
    
    return upc, isbn

# Create an Excel workbook
workbook = Workbook()
sheet = workbook.active
sheet.append(["Image File", "UPC", "ISBN"])

# Directory containing the images
image_dir = "images"

# Example usage
for filename in os.listdir(image_dir):
    if filename.endswith(".JPG") or filename.endswith(".png"):
        image_path = os.path.join(image_dir, filename)
        upc, isbn = extract_barcodes(image_path)
        
        # Append the barcode data to the Excel sheet
        sheet.append([filename, upc if upc else "", isbn if isbn else ""])

# Save the Excel file
workbook.save("barcode_data.xlsx")
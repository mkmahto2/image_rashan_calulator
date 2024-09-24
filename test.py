import pytesseract

# Set the path to tesseract.exe (replace with the correct path)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Now you can use pytesseract to process the image
from PIL import Image

image_path = 'product.png'
img = Image.open(image_path)
text = pytesseract.image_to_string(img)

print("Extracted text from image:")
print(text)


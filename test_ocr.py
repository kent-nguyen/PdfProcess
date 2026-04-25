import easyocr

IMAGE_PATH = r"Pages\2\Cells\row000_col000.png"

print(f"Testing OCR on: {IMAGE_PATH}")
reader = easyocr.Reader(["en"])
result = reader.readtext(IMAGE_PATH)
print("Raw result:", result)

text = " ".join(item[1] for item in result)
print("Text:", text)

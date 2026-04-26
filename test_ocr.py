import easyocr

IMAGE_PATH = r"Pages\2\Cells\row000_col000.png"

print(f"Đang kiểm tra OCR trên: {IMAGE_PATH}")
reader = easyocr.Reader(["en"])
result = reader.readtext(IMAGE_PATH)
print("Kết quả thô:", result)

text = " ".join(item[1] for item in result)
print("Văn bản:", text)

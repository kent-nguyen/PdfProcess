from pdf2image import convert_from_path
from PIL import ImageFilter, ImageOps
import os
import argparse
from libs.page_range import parse_pages

POPPLER_PATH = None  # e.g. r"C:\poppler\Library\bin" if not on system PATH


def convert_pages(pages_arg=None):
    if pages_arg:
        page_numbers = parse_pages(str(pages_arg))
        all_pages = convert_from_path(
            "Source.pdf", dpi=300,
            first_page=page_numbers[0],
            last_page=page_numbers[-1],
            poppler_path=POPPLER_PATH,
        )
        offset = page_numbers[0]
        pages_to_save = [(page_num, all_pages[page_num - offset]) for page_num in page_numbers]
    else:
        all_pages = convert_from_path("Source.pdf", dpi=300, poppler_path=POPPLER_PATH)
        pages_to_save = [(i + 1, page) for i, page in enumerate(all_pages)]

    for page_num, page in pages_to_save:
        folder = os.path.join("Pages", str(page_num))
        os.makedirs(folder, exist_ok=True)
        page = page.convert("L")                          # grayscale — B&W doc needs no colour
        page = ImageOps.autocontrast(page, cutoff=1)      # stretch histogram, ignore 1% outliers
        page = page.filter(ImageFilter.UnsharpMask(radius=1, percent=150, threshold=3))
        page.save(os.path.join(folder, f"{page_num}.png"), "PNG")
        print(f"Đã lưu trang {page_num}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert PDF pages to PNG images.")
    parser.add_argument(
        "--pages",
        type=str,
        default=None,
        help="Pages to convert, e.g. '1' or '1,3,5' or '2-5' or '1-3,7,10-12'. Omit to convert all pages.",
    )
    args = parser.parse_args()
    convert_pages(args.pages)

import pdfplumber
import os
import re
import win32com.client
from win32com.client import Dispatch


path = os.getcwd()


########################################################

pdf = r'C:\Users\Robert\Desktop\python\excel\zapis w Bazie\polisy\gen.pdf'

with pdfplumber.open(pdf) as policy:
    data = policy.pages[0].extract_text()
    words_separatly = re.compile(r"((?:(?<!'|\w)(?:\w-?'?)+(?<!-))|(?:(?<='|\w)(?:\w-?'?)+(?=')))")
    data = words_separatly.findall(data.lower())


def names_list(d):
    """Zwraca listę imion do programu"""
    with open(path + '\\imiona.txt') as content:
        all_names = content.read().split('\n')
    name = [f'{d[k + 1].title()} {v.title()}' for k, v in d.items() if v.title() in all_names]
    return ''.join(name[0])


# print(text_ocr)
print()
d = dict(enumerate(data))
print()
# print(d)

print(names_list(d))













##########################################################
# from PIL import Image
# from wand.image import Image as wi
# import pytesseract

# import io
# import re
#
#
# pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'
#

# pdf_path = r'C:\Users\Robert\Desktop\python\excel\zapis w Bazie\polisy\basen.jpg'
#
# """konwersja na .tiff do wszystkich dokumentów lub .png do HESTII"""
# pdf = wi(filename=pdf_path, resolution=200, format='raw')
#
# """docelowo 'pc_' in pdf_path"""
# p = 1 if 'pzu' in pdf_path else 2
#
#
# def ocr_text(pdf, ext, p):
#     pdfImage = pdf.convert(ext)
#     imgBlobs = []
#     for img in pdfImage.sequence[:p]: ### PZU ???
#         page = wi(image=img)
#         imgBlobs.append(page.make_blob(ext))
#
#     all_pages = []
#     for img in imgBlobs:
#         im = Image.open(io.BytesIO(img))
#         text_ocr = pytesseract.image_to_string(im, lang='eng')
#         words_separatly = re.compile(r"((?:(?<!'|\w)(?:\w-?'?)+(?<!-))|(?:(?<='|\w)(?:\w-?'?)+(?=')))")
#         data = words_separatly.findall(text_ocr.lower())
#         for i in data:
#             all_pages.append(i)
#
#         return all_pages, text_ocr
#
#
# try:
#     data, text_ocr = ocr_text(pdf, 'tiff', p)
# except:
#     data, text_ocr = ocr_text(pdf, 'png', p)





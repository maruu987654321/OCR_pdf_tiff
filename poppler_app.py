from pdf2image import convert_from_path
import time
import pytesseract
from PIL import Image
from PIL import Image, ImageSequence
import os
import cv2
from docx import Document
from docx.shared import Inches
import time
import numpy as np
import glob
import concurrent.futures
import re
from multiprocessing.pool import ThreadPool as Pool

Image.MAX_IMAGE_PIXELS = None
os.environ['OMP_THREAD_LIMIT'] = '1'
pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

start_time = time.time()

def remove_files(mypath):
    for root, dirs, files in os.walk(mypath):
        for file in files:
            os.remove(os.path.join(root, file))

def create_docx(text_list, path_result, name):
    document = Document()
    for j in text_list:
        p = document.add_paragraph(j)
    document.save(r"result_docx\{}.docx".format(name))

def convert_tiff_to_image(tif):
    im = Image.open(tif)
    index = 1
    for frame in ImageSequence.Iterator(im):
            frame.save(r"ghost_image_jpg\{}.jpg".format(index),  quality=600, compression='jpeg')
            index = index + 1

def ocr(img_path):
    result = []
    img = cv2.imread(img_path)
    text = pytesseract.image_to_string(img)
    text = text.replace('\n\n', '\n')
    out_file = re.sub(".jpg",".txt",img_path.split("\\")[-1])
    result.append(text)
    return out_file, result

def convert_tiff_to_image(tif):
    im = Image.open(tif)
    index = 1
    for frame in ImageSequence.Iterator(im):
        frame.save("image_jpg\\{}.jpg".format(index),  quality=400, compression='jpeg')
        index = index + 1

def to_jpg(path, path_to_save):
    pages = convert_from_path(path, 200, thread_count=8)  #this value you can change > (example 600)
    pdf_file = path.split('\\')[-1][:-4]
    print(pdf_file)
    for page in pages:
        page.save("{}\\{}-page{}.jpg".format(path_to_save, pdf_file,pages.index(page)), "JPEG")

def main_pdf(current_path, path_result, mypath_jpg, i):
    res_jpg = []
    res_pdf = []
    path = "{}".format(current_path) + "\\Deeds_sample\\{}".format(i)
    to_jpg(path, mypath_jpg)
    mypath_jpg = "{}".format(current_path) + "\\image_jpg" 
    all_process_jpg = os.listdir(mypath_jpg)
    for j in all_process_jpg:
        res_jpg.append(mypath_jpg + "\\{}".format(j))
    with concurrent.futures.ProcessPoolExecutor(max_workers=8) as executor: #multiprocessing
        for img_path,out_file in zip(res_jpg,executor.map(ocr,res_jpg)):
            res_pdf.append(out_file[1])
    print('Creating docx .....')
    name = i.split(".")[0]
    create_docx(res_pdf, path_result, name)
    remove_files(mypath_jpg)

def main_pdf_bad(current_path, path_result, mypath_jpg, item):
    res_jpg = []
    res_pdf = []
    to_jpg(item, mypath_jpg)
    mypath_jpg = "{}".format(current_path) + "\\image_jpg" 
    all_process_jpg = os.listdir(mypath_jpg)
    for j in all_process_jpg:
        res_jpg.append(mypath_jpg + "\\{}".format(j))
    with concurrent.futures.ProcessPoolExecutor(max_workers=8) as executor: #multiprocessing
        for img_path,out_file in zip(res_jpg,executor.map(ocr,res_jpg)):
            res_pdf.append(out_file[1])
    print('Creating docx .....')
    name = item.split("\\")[-1].split(".")[0]
    create_docx(res_pdf, path_result, name)
    remove_files(mypath_jpg)

def main_tif(current_path, path_result, mypath_jpg, i):
    res_jpg = []
    res_tif = []
    path = "{}".format(current_path) + "\\Deeds_sample\\{}".format(i)
    convert_tiff_to_image(path)
    mypath_jpg = "{}".format(current_path) + "\\image_jpg" 
    all_process_jpg = os.listdir(mypath_jpg)
    for j in all_process_jpg:
        res_jpg.append(mypath_jpg + "\\{}".format(j))
    with concurrent.futures.ProcessPoolExecutor(max_workers=8) as executor: #multiprocessing
        for img_path,out_file in zip(res_jpg,executor.map(ocr,res_jpg)):
            res_tif.append(out_file[1])
    name = i.split(".")[0]
    print('Creating docx .....')
    create_docx(res_tif, path_result, name)
    remove_files(mypath_jpg)

def process_bad_tif(path_to_save, name_tiff, i):
    tif_path = path_pdf_tif + "\{}".format(i)
    os.system(r"magick  -quiet {} -compress lzw {}\{}.pdf".format(tif_path, path_to_save, name_tiff))


if __name__ == '__main__':
    current_path = os.getcwd()
    path_pdf_tif = "{}".format(current_path) + "\\Deeds_sample"   #path to pdf and tif 
    if not os.path.exists('image_jpg'):
        os.makedirs('image_jpg')
    mypath_jpg = "{}".format(current_path) + "\\image_jpg" 
    if not os.path.exists('result_docx'):
        os.makedirs('result_docx')
    path_result = "{}".format(current_path) + "\\result_docx" 
    all_files = os.listdir(path_pdf_tif)
    all_process_jpg = os.listdir(mypath_jpg)
    if not os.path.exists('result_docx'):
        os.makedirs('result_docx')
    path_result = "{}".format(current_path) + "\\result_docx" 
    if not os.path.exists('bad_tif_to_pdf'):
        os.makedirs('bad_tif_to_pdf')
    path_to_save = "{}".format(current_path) + "\\bad_tif_to_pdf"
    for i in all_files:
       
        print(i)
        if "pdf" in i:
            main_pdf(current_path, path_result, mypath_jpg, i)
        if "tif" in i:
            try:
                tif_path = path_pdf_tif + "\\{}".format(i)
                main_tif(current_path, path_result, mypath_jpg, i)
            except OSError:
                name_tiff = i.split('.')[0]
                tif_path = path_pdf_tif + "\\{}".format(i)
                print('{} save in bad_tif_to_pdf, then process'.format(name_tiff))
                process_bad_tif(path_to_save, name_tiff, i)
       
    all_bad_file_tiff = os.listdir(path_to_save)
    for item in all_bad_file_tiff:
        item = path_to_save + "\\{}".format(item)
        result = []
        main_pdf_bad(current_path, path_result, mypath_jpg, item)
    remove_files(path_to_save)
    print(time.time() - start_time)
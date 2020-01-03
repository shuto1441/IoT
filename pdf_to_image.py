#-*- coding: utf8 -*-
from pdf2image import convert_from_path
import os
# import time
import glob
from PIL import Image, ImageChops
import openpyxl

def pdf_to_image(file):
    images = convert_from_path(file)
    i = 0
    for image in images:
        image.save('tournament{}.png'.format(i), 'png')
        
        img = Image.open("./tournament"+str(i)+".png")
        
        img = img.resize((1200, 1200), Image.LANCZOS)
        img.save("./tournament"+str(i)+".png")
    
        i += 1

# 余白を削除する関数
def cropImage(img): #引数は画像の相対パス

    # 周りの部分は強制的にトリミング
    w, h = img.size
    box = (w*0.05, h*0.05, w*0.95, h*0.95)
    img = img.crop(box)

    # 背景色画像を作成
    bg = Image.new("RGB", img.size, img.getpixel((0, 0)))
    # bg.show()

    # 背景色画像と元画像の差分画像を作成
    diff = ImageChops.difference(img, bg)
    # diff.show()

    # 背景色との境界を求めて画像を切り抜く
    croprange = diff.convert("RGB").getbbox()
    crop_img = img.crop(croprange)
    # crop_img.show()

    return crop_img

def pdf_to_image2(file):
    path="C:/Users/mech-user/Desktop/IoT/pdf/"+file+".pdf"
    save_path="C:/Users/mech-user/Desktop/IoT/photo/"
    images = convert_from_path(path)
    i = 0
    for image in images:
        if file=="管理表":
            image=cropImage(image)
            if i==0:
                image.save(save_path+'next.png', 'png')
            else:
                image.save(save_path+'court.png', 'png')
            
        else:
            wb=openpyxl.load_workbook("C:/Users/mech-user/Desktop/IoT/"+file+".xlsx")
            image=cropImage(image)
            image.save(save_path+'tournament{}.png'.format(wb.sheetnames[i][:2]), 'png')
        i += 1
    os.remove(path)
"""
images = convert_from_path("C:/Users/mech-user/Desktop/IoT/0527.pdf")
i = 0
for image in images:
    image.save('tournament{}.GIF'.format(i), 'GIF')
    i += 1
"""
#!/usr/bin/env python
# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment
from openpyxl import Workbook
import win32com.client
from pywintypes import com_error
import pathlib
import os
import pprint
import Excel_to_pdf
import pdf_to_image



def excel_to_png(file):
    Excel_to_pdf.main2(file)
    pdf_to_image.pdf_to_image2(file)

def main():
    excel_to_png("トーナメント")
    excel_to_png("管理表")

if __name__ == '__main__':
    main()


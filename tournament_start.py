#!/usr/bin/env python
# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment
from openpyxl import Workbook
import numpy as np
from openpyxl import utils
import copy
import re
import slackweb
import time

def main(filename,num):
    wb=openpyxl.load_workbook(filename)
    court=wb.worksheets[2]
    slack = slackweb.Slack(url="https://hooks.slack.com/services/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")
    slack.notify(text="-court"+str(num))
    time.sleep(5)

    for i in range(1,court.max_row+1):
        result=str(court['A'+str(i)].value)+","+court['B'+str(i)].value+","+str(court['C'+str(i)].value)+","+court['D'+str(i)].value+","+court['E'+str(i)].value+","+court['J'+str(i)].value+","+court['K'+str(i)].value
        slack.notify(text="-game"+result)
        time.sleep(5)

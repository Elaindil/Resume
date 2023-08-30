from PIL import Image
import imagehash

from pathlib import Path
import glob
import requests
from requests_html import HTMLSession, HTML
import time
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
startTime = datetime.now()
import os

directory = '.'
compareToMe= ''
originalimage = '100044252414b0'
newimage = '100044252414b2'
for filename in sorted(os.listdir(directory)):
    if 'b0' in filename:
        compareToMe = filename
        print (compareToMe)
        originalimage = compareToMe
        originalimagehash = imagehash.average_hash(Image.open(originalimage))
        print(originalimagehash)
    if 'b0' not in filename:
        newimage = filename
        newimagehash = imagehash.average_hash(Image.open(newimage))
        print(newimagehash)
        if originalimagehash == newimagehash:
            os.rename(newimage, newimage.replace("b", "дубликат"))
            
    
print(datetime.now() - startTime)  
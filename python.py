import pandas as pd
from docxtpl import DocxTemplate, RichText, InlineImage
import qrcode
import code128
import random
from docx.shared import Cm
from random import choice
from string import ascii_uppercase
from PIL import Image
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF, renderPM


def textEditor(string, fontsize):
    return RichText(text=string, size=fontsize)


def makeQrCode(id, url):
    img = qrcode.make(url)
    img.save(f"qr/qr{str(id)}.png")


def makeBarCode(id):
    with open(f"bar/bar{str(id)}.svg", "w") as f:
        f.write(code128.svg(id))
    drawing = svg2rlg(f"bar/bar{str(id)}.svg")
    renderPM.drawToFile(drawing, f"bar/bar{str(id)}.png", fmt="PNG")


def makeDoc(context, num, isLast):
    doc = DocxTemplate("front.docx")
    doc.render(context)
    if isLast:
        doc.save("Last.docx")
    else:
        doc.save(f"FRONT-{abs(num-9)+1}-{num+1}.docx")


def makeLastDoc(context, num, isLast, upic):
    doc = DocxTemplate("renderrear1.docx")
    if isLast:
        print("here in last")
        for i in range(0, num):
            doc.replace_media(f'images/qr{i+1}.png', f'qr/qr{upic[i]}.png')
            doc.replace_media(f'images/bar{i+1}.png', f'bar/bar{upic[i]}.png')
        doc.render(context)
        doc.save("Last.docx")
    else:
        for i in range(0, 10):
            print(f"Replacing bar {i+1}")
            doc.replace_media(f'images/bar{i+1}.png', f'bar/bar{upic[i]}.png')
            doc.replace_media(f'images/qr{i+1}.png', f'qr/qr{upic[i]}.png')
        doc.render(context)
        doc.save(f"REAR-{abs(num-9)+1}-{num+1}.docx")


def setData(context, contextForBack, upic, batch, url, address, lastupic,
            lastbatch, lasturl, lastaddress, mapCode, mapCodeLast):

    # making front
    for i in range(0, len(upic), 10):
        for j in range(0, 10):
            context[f"id{j+1}"] = textEditor(upic[i + j], 22)
        makeDoc(context, i, False)
    # making rear
    for i in range(0, len(upic), 10):
        for j in range(0, 10):
            makeBarCode(upic[i + j])
        for j in range(0, 10):
            makeQrCode(upic[i + j], url[i + j])
        for j in range(0, 10):
            contextForBack[f"batch{j+1}"] = textEditor(batch[i + j], 20)
            contextForBack[f"address{j+1}"] = textEditor(address[i + j], 15)
            contextForBack[f"map{j+1}"] = textEditor(mapCode[i+j], 26)
            # print(f"{i+1}")
            # print(contextForBack[f"map{i+1}"])
        makeLastDoc(contextForBack, i, False, upic)
    lastList = lasturl
    if lastList is not None:
        context = {
            'id1': '',
            'id2': '',
            'id3': '',
            'id4': '',
            'id5': '',
            'id6': '',
            'id7': '',
            'id8': '',
            'id9': '',
            'id10': ''
        }
        for i in range(0, len(lastList)):
            context[f"id{i+1}"] = textEditor(lastList[i], 22)
        makeDoc(context, 0, True)

        context = {
            "address1": '',
            "address2": '',
            "address3": '',
            "address4": '',
            "address5": '',
            "address6": '',
            "address7": '',
            "address8": '',
            "address9": '',
            "address10": '',
            "batch1": '',
            "batch2": '',
            "batch3": '',
            "batch4": '',
            "batch5": '',
            "batch6": '',
            "batch7": '',
            "batch8": '',
            "batch9": '',
            "batch10": '',
            "map1": '',
            "map2": '',
            "map3": '',
            "map4": '',
            "map5": '',
            "map6": '',
            "map7": '',
            "map8": '',
            "map9": '',
            "map10": '',
        }
        for i in range(0, len(lastupic)):
            context[f"address{i+1}"] = textEditor(lastaddress[i], 15)
            context[f"batch{i+1}"] = textEditor(lastbatch[i], 20)
            context[f"map{i+1}"] = textEditor(mapCodeLast[i], 26)
            makeBarCode(lastupic[i])
            makeQrCode(lastupic[i], lasturl[i])
        makeDoc(context, len(lastupic), True)
        # make Back Doc


df = pd.read_csv("csv.csv")
context = {
    'id1': '',
    'id2': '',
    'id3': '',
    'id4': '',
    'id5': '',
    'id6': '',
    'id7': '',
    'id8': '',
    'id9': '',
    'id10': ''
}
contextForBack = {
    "address1": '',
    "address2": '',
    "address3": '',
    "address4": '',
    "address5": '',
    "address6": '',
    "address7": '',
    "address8": '',
    "address9": '',
    "address10": '',
    "batch1": '',
    "batch2": '',
    "batch3": '',
    "batch4": '',
    "batch5": '',
    "batch6": '',
    "batch7": '',
    "batch8": '',
    "batch9": '',
    "batch10": '',
    'url1': '',
    'url2': '',
    'url3': '',
    'url4': '',
    'url5': '',
    'url6': '',
    'url7': '',
    'url8': '',
    'url9': '',
    'url10': '',
    'id1': '',
    'id2': '',
    'id3': '',
    'id4': '',
    'id5': '',
    'id6': '',
    'id7': '',
    'id8': '',
    'id9': '',
    'id10': '',
    "map1": '',
    "map2": '',
    "map3": '',
    "map4": '',
    "map5": '',
    "map6": '',
    "map7": '',
    "map8": '',
    "map9": '',
    "map10": '',
}
s = []
upic = []
batch = []
url = []
address = []
mapCode = []

for index, row in df.iterrows():
    upic.append(row["UpicNumber"])
    url.append(row["URL"])
    mapCode.append(row["MapCode"])
    batch.append(row["batch"])
    address.append(row["Address"])

forHowManyTimes = len(upic) // 10
leftElements = len(
    upic
) % 10  # always check if this variable is 0 or not. if 0, variable is of no use

if leftElements != 0:
    upicLast = upic[0:-leftElements]
    urlLast = url[0:-leftElements]
    batchLast = batch[0:-leftElements]
    addressList = address[0:-leftElements]
    tempList = s[-leftElements:]
    mapCodeLast = mapCode[0:-leftElements]
    setData(context, contextForBack, upic, batch, url, address, upicLast,
            batchLast, urlLast, addressList, mapCode, mapCodeLast)

else:
    urlLast = []
    batchLast = []
    addressList = []
    setData(context, contextForBack, upic, batch, url, address, [], [], [],
            [], mapCode, [])

# 4.68X 0.71
# Y1.67
# 1.89
import pandas as pd
from docxtpl import DocxTemplate, RichText, InlineImage
import qrcode
import code128
import random
from docx.shared import Cm, Mm
from random import choice
from string import ascii_uppercase
from PIL import Image
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF, renderPM

doc = DocxTemplate("check.docx")

# context = {
#     'img1': '',
#     'img2': '',
#     'img3': '',
#     'img4': '',
#     'img5': '',
#     'img6': '',
# }

# context["img1"] = InlineImage(doc,
#                               image_descriptor="img1.png",
#                               width=Mm(25),
#                               height=Mm(25))
# context["img2"] = InlineImage(doc,
#                               image_descriptor="img2.jpeg",
#                               width=Mm(25),
#                               height=Mm(25))
# context["img3"] = InlineImage(doc,
#                               image_descriptor="img3.jpeg",
#                               width=Mm(25),
#                               height=Mm(25))

# context["img4"] = InlineImage(doc,
#                               image_descriptor="img3.jpeg",
#                               width=Mm(25),
#                               height=Mm(25))

# context["img5"] = InlineImage(doc,
#                               image_descriptor="img2.jpeg",
#                               width=Mm(25),
#                               height=Mm(25))

# context["img6"] = InlineImage(doc,
#                               image_descriptor="img1.png",
#                               width=Mm(25),
#                               height=Mm(25))
# doc.render(context)
doc.replace_media('img1.png', '1.png')
doc.replace_media('img2.jpeg', '2.png')
doc.replace_media('img3.jpeg', '3.png')
doc.save("check-new.docx")
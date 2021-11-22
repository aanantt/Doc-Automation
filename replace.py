from docxtpl import DocxTemplate, RichText, InlineImage
from docx.shared import Cm,Mm

doc = DocxTemplate("rear.docx")

context = {
    'bar1': '',
    'bar2': '',
    'bar3': '',
    'bar4': '',
    'bar5': '',
    'bar6': '',
    'bar7': '',
    'bar8': '',
    'bar9': '',
    'bar10': '',
    'qr1': '',
    'qr2': '',
    'qr3': '',
    'qr4': '',
    'qr5': '',
    'qr6': '',
    'qr7': '',
    'qr8': '',
    'qr9': '',
    'qr10': '',
}

for i in range(0, 10):
    context[f"qr{i+1}"] = InlineImage(doc,
                                      image_descriptor=f'images/qr{i+1}.png',
                                      width=Mm(25),
                                      height=Mm(25))
    context[f"bar{i+1}"] = InlineImage(doc,
                                      image_descriptor=f'images/bar{i+1}.png',
                                      width=Mm(30),
                                      height=Mm(10))

print(context)
doc.render(context)
doc.save("renderrear1.docx")


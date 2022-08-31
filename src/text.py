"""
Redefinition of the 'text' property 
of docx allowing for hyperlinks

Author: roydesbois
Source: https://github.com/python-openxml/python-docx/issues/85#issuecomment-917134257
Date:   9/10/2021
"""

from docx.text.paragraph import Paragraph
import re

from loader import LINK_BEG, LINK_END

Paragraph.text = property(lambda self: GetParagraphText(self))

def GetTag(element):
    return "%s:%s" % (element.prefix, re.match("{.*}(.*)", element.tag).group(1))

def GetParagraphText(paragraph):
    text = ''
    runCount = 0
    for child in paragraph._p:
        tag = GetTag(child)
        if tag == "w:r":
            text += paragraph.runs[runCount].text
            runCount += 1
        if tag == "w:hyperlink":
            for subChild in child:
                if GetTag(subChild) == "w:r":
                    text += f'{LINK_BEG}{subChild.text}{LINK_END}'
    return text
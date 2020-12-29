
import docx 
from docx.shared import Pt

doc = docx.Document('f.docx') 
print(len(doc.paragraphs))
c = doc.paragraphs[0].style.font.color.rgb


print(c)




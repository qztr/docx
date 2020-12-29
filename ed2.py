
import docx 
from docx.shared import Pt

doc = docx.Document('hand.docx') 
print(len(doc.paragraphs))
par1 = doc.paragraphs[0]
t = []
sizes = []
names = []
for run in par1.runs:
  t.append(run.text)

  if run.font.size is not None: 
   sizes.append(str(run.font.size)) 
    
  if run.font.name is not None: 
   names.append(str(run.font.name)) 
   
#for i in range(len(t)): 
# print(t[i])
print(' '.join(t)) 
print(' '.join(sizes)) 
print(' '.join(names)) 




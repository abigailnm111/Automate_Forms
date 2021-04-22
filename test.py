
"""
Created on Tue Apr 13 10:24:12 2021

@author: panda
"""

from datetime import date
import re
from docx import Document
import docx2txt

template= docx2txt.process('Pre6-letter-template.docx')
new_doc=Document('Letterhead.docx')

last_name="George"
first_name="Regina"

new_doc.save(last_name+"_"+first_name+'.docx')
date= date.today()
title_code="1630"
start= "7/1/2021"
end= "6/30/2022"
percentage="100%"
annual= "60,000"
monthly="5,000"
courses= "3 EC 3 ksldjfskjfsldkfj"

data={"<Date>":date ,"<Name>":first_name+" "+last_name, 
      "<Title Code>": title_code, "<start>": start, "<end>": end, 
      "<percentage>":percentage, "<annual>":annual, "<monthly>": monthly,
      "<List of Classes& Course Titles>": courses}


template=re.sub("<First Name><Last Name>", first_name+" "+last_name, template)
print(template)
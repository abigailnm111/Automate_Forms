
"""
Created on Tue Apr 13 10:24:12 2021

@author: panda
"""

from datetime import date
import re
from docx import Document
import docx2txt

def write_pre_offer_letter(lecturer):
    template= docx2txt.process('Pre6-letter-template.docx')
    new_doc=Document('Letterhead.docx')
    
    
    
    today= date.today()
    
    letter_date= today.strftime("%B %d, %Y")
    
    
    
    all_courses=lecturer.F_courses+lecturer.W_courses+lecturer.S_courses
    courses=" ".join(str(course)+", " for course in all_courses)
    job_code=" ".join(str(code) for code in lecturer.job_code )
    
    annual= " ".join(str(annual)+", " for annual in lecturer.annual)
    monthly= " ".join(str(monthly)+', ' for monthly in lecturer.monthly)
    i=0
    dates=' '
    
    for s_date in lecturer.start:
        start_end= lecturer.start[i]+" - "+lecturer.end[i]
        
        
        
        if i== 1:
            start_end= ", "+start_end
            
       dates= dates+start_end
       
        
        i+=1
   
    
    data={"<Date>":str(letter_date) ,"<Name>":lecturer.first_name+" "+lecturer.last_name, 
          "<Title Code>": job_code, "<start to end>": dates, 
          "<percentage>":lecturer.percentage, "<annual>":annual, "<monthly>": monthly,
          "<List of Classes & Course Titles>": courses, "\n\n": '\n'}
    
    for key in data:
        template=re.sub(key, data[key] , template)
    
    letter=new_doc.add_paragraph(template)
    letter_format= letter.paragraph_format
    letter_format.line_spacing=1.0
    new_doc.save(lecturer.last_name+"_"+lecturer.first_name+'.docx')
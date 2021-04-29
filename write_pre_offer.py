
"""
Created on Tue Apr 13 10:24:12 2021

@author: panda
"""

from datetime import date
import re
from docx import Document
import docx2txt

def format_course(course):
    for c in course:
            title= course.split(' ',1)
            if ("course" in title[1]) or ("Course" in title[1]):
                course_format= course
            else:   
                course_format= title[0]+" section(s): "+title[1]
            
    return course_format   


def format_courses_by_quarter(lecturer):
    all_courses=[]
    if lecturer.F_courses!=[]:
        c= "".join(str(format_course(course)+",\n\t\t\t\t") for course in lecturer.F_courses)
        courses=c[:-7]+c[-5:]
        all_courses.append("Fall: "+courses)
        
    if lecturer.W_courses!=[]:
      
        c= "".join(str(format_course(course)+",\n\t\t\t\t") for course in lecturer.W_courses)
        courses=c[:-7]+c[-5:]
        all_courses.append("Winter: "+c)
        
    if lecturer.S_courses!=[]:
        
        c= "".join(str(format_course(course)+",\n\t\t\t\t") for course in lecturer.S_courses)
        courses=c[:-7]+c[-5:]
        all_courses.append("Spring: "+c)

    courses= "".join(str(quarter) for quarter in all_courses)
    return courses

def write_pre_offer_letter(lecturer):
    template= docx2txt.process('Pre6-letter-template.docx')
    new_doc=Document('Letterhead.docx')
    
    
    
    today= date.today()
    
    letter_date= today.strftime("%B %d, %Y")
    courses=format_courses_by_quarter(lecturer)
    
    job_code=str(lecturer.job_code[0])
    
    
    i=0
    dates=' '
    
    
    
    
        
    
    for s_date in lecturer.start:
       # if lecturer.job_code[i]=="1630" or lecturer.job_code[i]== "1632":
            start_end= lecturer.start[i]+" - "+lecturer.end[i]
            annual=str(lecturer.annual[0])
            monthly=str(lecturer.monthly[0])
            
            if i== 1:
                start_end= ", "+start_end
                if len(lecturer.annual)>1:
                    if lecturer.annual[1]== "ATTENTION":
                        annual= annual+", "+lecturer.start[i]+': '+lecturer.annual[1]  
                        monthly=monthly+", "+lecturer.start[i]+': '+lecturer.monthly[1]
                
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
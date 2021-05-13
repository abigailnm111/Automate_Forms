#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun May  2 21:33:27 2021

@author: panda
"""
from PyPDF2.generic import NameObject, BooleanObject, TextStringObject, IndirectObject, DictionaryObject, ArrayObject,RectangleObject
from PyPDF2 import PdfFileReader, PdfFileWriter
from datetime import date

today= date.today()
present_date=today.strftime("%B %d, %Y")

def course_formatting(courses, qtr):
    course_dict={}
    for course in courses:
        for c in course:
                title= course.split(' ',2)
                if ("course" in title[1]) or ("Course" in title[1]):
                    sections=int(title[0])
                    c_num= ''
                    c_title=title[1]+" "+title[2]
                else:   
                    sections=int(title[0])
                    c_num=title[1]
                    c_title=title[2]
        
        i=1
        for s in range(0,sections):
            
            course_dict[qtr+'_iwc'+str(i)]="1"
            course_dict[qtr+'_course_no'+str(i)]=c_num
            course_dict[qtr+'_desc'+str(i)]=c_title
            i+=1 
    return course_dict

def quarter_info(lecturer):
    course_dict={}
    percentage_raw=int(lecturer.percentage[:-1])/100
    if lecturer.F_courses!= []:
       
       course_dict['fall_percent_p2']=percentage_raw
       courses=course_formatting(lecturer.F_courses,"fall")
       course_dict.update(courses)
       
    if lecturer.W_courses!=[]:
        course_dict['wtr_percent_p2']= percentage_raw
        courses=course_formatting(lecturer.W_courses, 'wtr')
        course_dict.update(courses)
    if lecturer.S_courses!=[]:
        course_dict['spr_percent_p2']=percentage_raw
        courses=course_formatting(lecturer.S_courses, 'spr')
        course_dict.update(courses)
            
    return course_dict
       
    
def form_dicts(lecturer,AY,HR,dept):
    course_dict=quarter_info(lecturer) 
    if HR=="new":
        present_title=""
        present_salary=""
    else:
        present_title=lecturer.title[0]
        present_salary="{:,}".format(lecturer.hr_salary)
        
    course_dict=quarter_info(lecturer)  
    
    if lecturer.job_code==[1630]:
        title_id=1
        title_text='1630, LECTURER 9/12'
    elif lecturer.job_code==[1632]:
        title_id=2
        title_text="1632, LECTURER 9/9"
    elif lecturer.job_code==[1640]:
        title_id=3
        title_text="1640, SENIOR LECTURER 9/12"
    elif lecturer.job_code== [1642]:
        title_id=4
        title_text='1642, SENIOR LECTURER 9/9'
    elif lecturer.job_code==[1631]:
        title_id=5
        title_text="1631, CONT LECTURER 9/12"
    else:
        title_id=6
        title_text="1633, CONT LECTURER 9/9"
        

    employee_dict={
            
            'name_p1':lecturer.first_name+" "+lecturer.last_name,
            'dept_p1': dept.name,
            'accrued_qtr':lecturer.quarters, #History record quarter count
            'present_title_p1':present_title,
            
            'present_salary':present_salary,
            'proposed_salary_p1':"{:,}".format(lecturer.annual[0]),
            
            'present_percent_p1':lecturer.hr_percentage,## not working
            'proposed_percent_p1':lecturer.percentage,
            'present_add_comment_p1':'',
            'begin_date_p1':lecturer.start,
            'end_date_p1':lecturer.end,
            # bio_data_note
            # degree1_p1
            # date1_p1
            # inst1_p1
            # degree2_p1
            # date2_p1
            # inst2_p1
            # degree3_p1
            # date3_p1
            # inst3_p1
            # degree4_p1
            # date4_p1
            # inst4_p1
            
            # other_state_inst_p2
            
            # visa_current
            'uid_p2':'12345678',
             'dept_code_p2': dept.code,
            'mo_salary_rate_p2': "{:,.2f}".format(lecturer.monthly[0]),
            
         
            'avg_percent_p2':lecturer.percentage,
            'fund_source_p2':dept.FAU,
            # phd_cert
            # yes
            # no
            # name_inst_p2
            # other_inst_percent
            # type_of_visa
            'dept_contact_name':dept.contact_name,
            'dept_ext':dept.contact_ext,
            'dept_contact_date':present_date
            
            # add_comments_p2
            
            # visa_beg date
            # visa_end date
            # JPF_EXR
            
                
            }
    
    employee_dict.update(course_dict)
    radio_dict={
        'type_increase': '/4th_year',
        'type_review_p1':'/reappt'
                
        }

    choice_dict={
        'proposed_title_p1':[title_id,title_text],
        }
    
    return employee_dict, radio_dict, choice_dict
def write_form(lecturer,employee_dict, radio_dict, choice_dict):

    output=PdfFileWriter()
    template=PdfFileReader(open("1_Pre6_form.pdf", 'rb'))

    output.cloneReaderDocumentRoot(template)
    output._root_object["/AcroForm"][NameObject("/NeedAppearances")]=BooleanObject(True)
    
    
    for i in [0,1]:
        
        output.updatePageFormFieldValues(template.getPage(i), employee_dict)
        
        
        #Checkboxes and drop downs:From PyPDF Library for updatePageFormFieldValues but edited for NameObject as value 
        
        page=template.getPage(i)
        
        for j in range(0, len(page['/Annots'])):
                writer_annot = page['/Annots'][j].getObject()   
                
                #dropdowns:changes "I" to index of option chosen ex: second option on list is "1"
                #"V" is the text of the field. Both V and I must be updated
                for field in choice_dict:
                    if writer_annot.get("/T")==field:
                        writer_annot.update({
                                NameObject("/I"):NameObject(choice_dict[field][0]),
                                NameObject("/V"):TextStringObject(choice_dict[field][1])
                                })
                        
                #checkboxes on this form are kids of a parent object.
                #accesses parent of object to get NameID
                for field in radio_dict:
                    if "/Parent" in writer_annot:
                        if writer_annot["/Parent"].get("/T") == field:
                            writer_annot.update({
                                NameObject("/V"):NameObject(radio_dict[field]),
                                NameObject("/AS"):NameObject(radio_dict[field])})
 
    outputStream=open(lecturer.last_name+"."+lecturer.first_name+"_form.pdf", "wb")
    output.write(outputStream)
   

def fill_form(lecturer,AY,HR,dept):
    employee_dict,radio_dict, choice_dict=form_dicts(lecturer,AY,HR,dept)
    write_form(lecturer, employee_dict, radio_dict, choice_dict)
    
def test():
   # test=PdfFileReader(open("NewPDF.pdf", 'rb'))
   # print(test.getPage(0)['/Annots'].getObject())
    output=PdfFileWriter()
    template=PdfFileReader(open("1_Pre6_form.pdf", 'rb'))

    
    #fields=template.getFields()
    #print(fields)
    print(template.getPage(0)['/Annots'])
    page=template.getPage(0)
    i =len(page['/Annots'])
    obj=output._addObject(DictionaryObject({ NameObject('/DA'):TextStringObject('/Helv 10 Tf 0 g'),
                 NameObject('/Subtype'):NameObject('/FreeText'),
                 NameObject('/Rect'):RectangleObject([378.481,426.012, 450.601, 426.892 ]),
                 NameObject('/Type'):NameObject('/Annot'), 
                 NameObject('/Contents'):TextStringObject('4/1/2020')} ))
    page['/Annots'].append(obj)
        
    print(template.getPage(0))
    #print(output._root_object)
    #template.getPage(0)['/Annots']
    
    output.cloneReaderDocumentRoot(template)
    output._root_object["/AcroForm"][NameObject("/NeedAppearances")]=BooleanObject(True)
    outputStream=open("NewPDF.pdf", "wb")
    output.write(outputStream)
test()

#begin date RECT [378.481, 427.012, 450.601, 438.892]
#end date RECT[465.361, 427.012, 537.361, 438.892]

#begin date 2[378.481,527.012, 450.601, 538.892 ]
#end date 2[465.361,527.012, 537.361,538.892]

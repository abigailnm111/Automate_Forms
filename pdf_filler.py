#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun May  2 21:33:27 2021

@author: panda
"""
from PyPDF2.generic import NameObject, BooleanObject, TextStringObject,  DictionaryObject, ArrayObject,RectangleObject, FloatObject
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
        type_review="/appt"
    else:
        present_title=lecturer.title[0]
        present_salary="{:,}".format(lecturer.hr_salary)
        type_review="/reappt"
    if '10th Quarter Increase' in lecturer.action_type:
        type_increase="/4th_year"
    else:
        type_increase= '/Off'
    try:
        prop_annual="{:,}".format(lecturer.annual[0])
        prop_mo="{:,.2f}".format(lecturer.monthly[0])
    except:
        prop_annual=lecturer.annual[0]
        prop_mo=lecturer.monthly[0]
    if len(lecturer.end)==2 and lecturer.break_service==False:
        end_date=lecturer.end[1]
    else:
        end_date=lecturer.end[0]
        
    course_dict=quarter_info(lecturer)  
    
    if lecturer.job_code==[1630]:
        title_id=1
        title_text='1630, LECTURER 9/12'
    elif lecturer.job_code==[1632]:
        title_id=2
        title_text="1632, LECTURER 9/9"
    elif lecturer.job_code==[1640]:
        title_id=3
        title_text="1641,SR CONTINUING LECTURER 9/12"
    elif lecturer.job_code== [1642]:
        title_id=4
        title_text='1643,SR CONTINUING LECTURER 9/9'
    elif lecturer.job_code==[1631]:
        title_id=1
        title_text="1631,CONTINUING LECTURER 9/12"
    else:
        title_id=2
        title_text="1633,CONTINUING LECTURER 9/9"
        
#basic form fields
    employee_dict={
            #page1
            'name_p1':lecturer.last_name+", "+lecturer.first_name,
            'dept_p1': dept.name,
            'accrued_qtr':lecturer.HRquarters, 
            'present_title_p1':present_title,
            
            'present_salary':present_salary,
            'proposed_salary_p1':prop_annual,
            
            'present_percent_p1':lecturer.hr_percentage,
            'proposed_percent_p1':lecturer.percentage,
            'present_add_comment_p1':'',
            'begin_date_p1':lecturer.start[0],
            'end_date_p1':end_date,
            #page2
            'uid_p2':lecturer.UID,
            'dept_code_p2': dept.code,
            'mo_salary_rate_p2':prop_mo,
            
         
            'avg_percent_p2':lecturer.percentage,
            'fund_source_p2':dept.FAU,

            'dept_contact_name':dept.contact_name,
            'dept_ext':dept.contact_ext,
            'dept_contact_date':present_date,
            'dept_approval_typed':dept.approver,
   
            }
    #adds course information to dictionary
    employee_dict.update(course_dict)
    
    #dictionary for radio buttons/checkboxes
    radio_dict={
        'type_increase': type_increase,#'/4th_year' or '/merit' or '/Off'
        'type_review_p1':type_review #'/appt' or '/reappt'
                
        }
    #dictionary for dropdowns
    choice_dict={
        'proposed_title_p1':[title_id,title_text],
        }
    #Continuing lecturer form has some additional fields and radio buttons change name
    if lecturer.status== ['sr'] or lecturer.status== ['cont']:
        
        choice_dict['present_title_p1']=[title_id, title_text]
        radio_dict['pre_six_year_reappointment']="/Yes"
        radio_dict['annual_notice_p2']="/Yes"
        
    
    return employee_dict, radio_dict, choice_dict

#accounts for multiple start dates and salaries by adding a comment box next to original field form
def add_comment(output,page,text, rectangle):
    obj=output._addObject(DictionaryObject({ NameObject('/DA'):TextStringObject(' /Helv 10 Tf'),
                     NameObject('/Subtype'):NameObject('/FreeText'),
                     NameObject('/Rect'):RectangleObject(rectangle),
                     NameObject('/Type'):NameObject('/Annot'), 
                     NameObject('/Contents'):TextStringObject(text),
                     NameObject('/C'):ArrayObject([FloatObject(1),FloatObject(1),FloatObject(1)]),
                     
                     } ))
    page['/Annots'].append(obj)
 
#crates form and fills out the fields using dictionaries    
def write_form(lecturer,employee_dict, radio_dict, choice_dict):

    output=PdfFileWriter()
    if lecturer.job_code[0]==1630 or lecturer.job_code[0]==1632:
        template=PdfFileReader(open("Pre6_Data_Summary__Redacted.pdf", 'rb'))
        filename= lecturer.last_name+"."+lecturer.first_name+"_form.pdf"
    else:
        template=PdfFileReader(open("CL_Data_Summary_form_Redacted.pdf", 'rb'))
        filename= lecturer.last_name+"."+lecturer.first_name+"_Cont_form.pdf"
    output.cloneReaderDocumentRoot(template)
    output._root_object["/AcroForm"][NameObject("/NeedAppearances")]=BooleanObject(True)
    
    for i in [0,1]:
        
        output.updatePageFormFieldValues(template.getPage(i), employee_dict)

        page=template.getPage(i)
 
        #Checkboxes and drop downs:From PyPDF Library for updatePageFormFieldValues but edited for NameObject as value
        for j in range(0, len(page['/Annots'])):
                writer_annot = page['/Annots'][j].getObject()   
                
                #dropdowns:changes "I" to index of option chosen ex: second option on list is "1"
                #"V" is the text of the field. Both V and I must be updated
                for field in choice_dict:
                  #  print(lecturer.last_name, field)
                    if writer_annot.get("/T")==field:
                        
                        writer_annot.update({
                                NameObject("/I"):NameObject(choice_dict[field][0]),
                                NameObject("/V"):TextStringObject(choice_dict[field][1])
                                })
                        
                #checkboxes on pre6 form are kids of a parent object.
                #accesses parent of object to get NameID
                #checkboxes on cont form are accesible by "/T" alone
                for field in radio_dict:
                    if "/Parent" in writer_annot:
                        if writer_annot["/Parent"].get("/T") == field:
                            
                            writer_annot.update({
                                NameObject("/V"):NameObject(radio_dict[field]),
                                NameObject("/AS"):NameObject(radio_dict[field])})
                    elif writer_annot.get("/T")==field:
                         writer_annot.update({
                            NameObject("/V"):NameObject(radio_dict[field]),
                            NameObject("/AS"):NameObject(radio_dict[field])})
        if i==0:
            # if there are two start dates and it's the first page add second set of dates to the proposed dates
            if len(lecturer.start)==2 and (lecturer.break_service==True): 
                 start_end_2= lecturer.start[1]+"-"+lecturer.end[1]
                 add_comment(output,page, start_end_2, [379.051, 405.913, 536.515, 424.313])
            #if they are eligible for a raise in the middle of the year add a line for a second monthly/annual(pg1 and pg2)
            if len(lecturer.annual)==2:
                add_comment(output,page,lecturer.start[1]+": "+lecturer.annual[1],[457.783, 465.165, 582.78, 483.565])
        if i==1:
            if len(lecturer.monthly)==2:
                   add_comment(output, page, lecturer.start[1]+": "+lecturer.monthly[1], [440.738, 679.446, 548.738, 697.846])
 
    outputStream=open(filename, "wb")
    output.write(outputStream)
   
#main function to call from main file 
def fill_form(lecturer,AY,HR,dept):
    employee_dict,radio_dict, choice_dict=form_dicts(lecturer,AY,HR,dept)
    write_form(lecturer, employee_dict, radio_dict, choice_dict)
    



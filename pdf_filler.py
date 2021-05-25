#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun May  2 21:33:27 2021

@author: panda
"""
from PyPDF2.generic import NameObject, BooleanObject, TextStringObject, IndirectObject, DictionaryObject, ArrayObject,RectangleObject, FloatObject
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
        

    employee_dict={
            
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
            #'uid_p2':'12345678',
            'dept_code_p2': dept.code,
            'mo_salary_rate_p2':prop_mo,
            
         
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
        'type_increase': type_increase,#'/4th_year' or '/merit' or '/Off'
        'type_review_p1':type_review #'/appt' or '/reappt'
                
        }

    choice_dict={
        'proposed_title_p1':[title_id,title_text],
        }
    if lecturer.status== ['sr'] or lecturer.status== ['cont']:
        
        choice_dict['present_title_p1']=[title_id, title_text]
        radio_dict['annual_notice_p2']="/Yes"
        
    
    return employee_dict, radio_dict, choice_dict

def add_comment(output,page,text, rectangle):
    obj=output._addObject(DictionaryObject({ NameObject('/DA'):TextStringObject(' /Helv 10 Tf'),
                     NameObject('/Subtype'):NameObject('/FreeText'),
                     NameObject('/Rect'):RectangleObject(rectangle),
                     NameObject('/Type'):NameObject('/Annot'), 
                     NameObject('/Contents'):TextStringObject(text),
                     NameObject('/C'):ArrayObject([FloatObject(1),FloatObject(1),FloatObject(1)]),
                     
                     } ))
    page['/Annots'].append(obj)
    
def write_form(lecturer,employee_dict, radio_dict, choice_dict):

    output=PdfFileWriter()
    if lecturer.job_code==[1630] or lecturer.job_code==[1632]:
        template=PdfFileReader(open("1_Pre6_form.pdf", 'rb'))
    else:
        template=PdfFileReader(open("CL_Data_Summary_form.pdf", 'rb'))

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
                  #  print(lecturer.last_name, field)
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
                            print(lecturer.last_name, field)
                            writer_annot.update({
                                NameObject("/V"):NameObject(radio_dict[field]),
                                NameObject("/AS"):NameObject(radio_dict[field])})
                    elif writer_annot.get("/T")==field:
                         print (lecturer.last_name, field)
                         writer_annot.update({
                            NameObject("/V"):NameObject(radio_dict[field]),
                            NameObject("/AS"):NameObject(radio_dict[field])})
        if i==0:
            if len(lecturer.start)==2 and (lecturer.break_service==True): # if there are two start dates and it's the first page
                 start_end_2= lecturer.start[1]+"-"+lecturer.end[1]
                 add_comment(output,page, start_end_2, [379.051, 405.913, 536.515, 424.313])
            if len(lecturer.annual)==2:
                add_comment(output,page,lecturer.start[1]+": "+lecturer.annual[1],[457.783, 465.165, 582.78, 483.565])
        if i==1:
            if len(lecturer.monthly)==2:
                   add_comment(output, page, lecturer.start[1]+": "+lecturer.monthly[1], [440.738, 679.446, 548.738, 697.846])
 
    outputStream=open(lecturer.last_name+"."+lecturer.first_name+"_form.pdf", "wb")
    output.write(outputStream)
   

def fill_form(lecturer,AY,HR,dept):
    employee_dict,radio_dict, choice_dict=form_dicts(lecturer,AY,HR,dept)
    write_form(lecturer, employee_dict, radio_dict, choice_dict)
    
def test():
    
    output=PdfFileWriter()
    template=PdfFileReader(open("Heron.Cady_form.pdf", 'rb'))
    fields=template.getFields()
    #for field in fields:
        #print(field)
    page=template.getPage(0)
    for j in range(0, len(page['/Annots'])):
                writer_annot = page['/Annots'][j].getObject()
                print(writer_annot)
    
     
    
    #print(output._root_object)
    #print(template.getPage(1)['/Annots'][74].getObject())
    
    output.cloneReaderDocumentRoot(template)
    
    output._root_object["/AcroForm"][NameObject("/NeedAppearances")]=BooleanObject(True)
    outputStream=open("NewPDF.pdf", "wb")
    output.write(outputStream)
test()

#{'/AP': {'/D': {'/Off': IndirectObject(22, 0), '/Yes': IndirectObject(21, 0)}, '/N': {'/Off': IndirectObject(24, 0), '/Yes': IndirectObject(23, 0)}}, '/AS': '/Off', '/BS': {'/S': '/I', '/W': 1}, '/DA': '/ZaDb 14 Tf 0 g', '/F': 4, '/FT': '/Btn', '/MK': {'/BC': [0], '/BG': [1], '/CA': '4'}, '/P': IndirectObject(1, 0), '/Rect': [32.2339, 664.372, 48.3508, 679.617], '/Subtype': '/Widget', '/T': 'annual_notice_p2', '/Type': '/Annot'}

# uid_p2
# mo_salary_rate_p2 same
# dept_code_p2 same
# assign_revision_p2
# assign_rev_percent_p2
# fall_percent_p2 same
# avg_percent_p2 same
#

# other_inst_percent
# yes
# other_state_inst_p2
# no
# name_inst_p2
# add_comments_p2
# dept_contact_name
# dept_approval_typed
# dept_ext
# dept_contact_date
# dept_approval_date




# present_salary
# proposed_salary_p1
# present_percent_p1
# proposed_percent_p1
# present_add_comment_p1
# begin_date_p1
# end_date_p1
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
# dean_final_p1
# dean_action_p1
# dean_date_p1
# dean_name_p1

# pre_six_year_appointment
# pre_six_year_merit
# pre_six_year_reappointment
# pre_six_year_4th_year_increase
# revision
# dept_approval
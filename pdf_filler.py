#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun May  2 21:33:27 2021

@author: panda
"""

import PyPDF2
def fill_form(lecturer,AY,HR,dept):
    
    if HR=="new":
        present_title=""
        present_salary=""
    else:
        present_title=lecturer.title[0]
        present_salary="{:,}".format(lecturer.hr_salary)
        
    print(lecturer.last_name, lecturer.hr_percentage)
    employee_dict={
            
            'name_p1':lecturer.first_name+" "+lecturer.last_name,
            'dept_p1': dept.name,
            'accrued_qtr':lecturer.quarters, #History record quarter count
            'present_title_p1':present_title,
            #**proposed_title_p1
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
            
            # fall_percent_p2
            # fall_iwc1
            # fall_course_no1
            # fall_desc1
            # fall_iwc2
            # fall_course_no2
            # fall_desc2
            # fall_iwc3
            # fall_course_no3
            # fall_desc3
            # fall_iwc4
            # fall_course_no4
            # fall_desc4
            # wtr_percent_p2
            # wtr_iwc1
            # wtr_course_no1
            # wtr_desc1
            # wtr_iwc2
            # wtr_course_no2
            # wtr_desc2
            # wtr_iwc3
            # wtr_course_no3
            # wtr_desc3
            # wtr_iwc4
            # wtr_course_no4
            # wtr_desc4
            # spr_percent_p2
            # spr_iwc1
            # spr_course_no1
            # spr_desc1
            # spr_iwc2
            # spr_course_no2
            # spr_desc2
            # spr_iwc3
            # spr_course_no3
            # spr_desc3
            # spr_iwc4
            # spr_course_no4
            # spr_desc4
            # avg_percent_p2
            'fund_source_p2':dept.FAU,
            # phd_cert
            # yes
            # no
            # name_inst_p2
            # other_inst_percent
            # type_of_visa
            'dept_contact_name':dept.contact_name,
            'dept_ext':dept.contact_ext,
            # dept_contact_date
            
            # add_comments_p2
            
            # visa_beg date
            # visa_end date
            # JPF_EXR
            
                
            }
    
    radio_dict={
        'type_increase': '/4th_year',
        'type_review_p1':'/reappt'
                
        }


    



    output=PyPDF2.PdfFileWriter()
    template=PyPDF2.PdfFileReader(open("1_Pre6_form.pdf", 'rb'))
    fields=template.getFields()
    #for field in fields:
        
       # print(field)
    
    output.addPage(template.getPage(0))
    output.addPage(template.getPage(1))
    
    for i in [0,1]:
        output.updatePageFormFieldValues(output.getPage(i), employee_dict)
        
        
        #From PyPDF Library for updatePageFormFieldValues but edited for NameObject as value
        #accesses parent of object to get NameID
        page=output.getPage(i)
        
        for j in range(0, len(page['/Annots'])):
                writer_annot = page['/Annots'][j].getObject()    
                for field in radio_dict:
                    if "/Parent" in writer_annot:
                        if writer_annot["/Parent"].get("/T") == field:
                            writer_annot.update({
                                PyPDF2.generic.NameObject("/V"):PyPDF2.generic.NameObject(radio_dict[field]),
                                PyPDF2.generic.NameObject("/AS"):PyPDF2.generic.NameObject(radio_dict[field])})
        
    #print(output.getPage(1)['/Annots'][0].getObject())
    
  
    
    outputStream=open(lecturer.last_name+"."+lecturer.first_name+"_form.pdf", "wb")
    output.write(outputStream)
   




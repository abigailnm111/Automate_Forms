#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun May  2 21:33:27 2021

@author: panda
"""

import pdfrw
import PyPDF2

employee_dict={
        
        'name_p1':' Regina George',
        'dept_p1': "Mathletes",
        'accrued_qtr':'10', #History record quarter count
        'type_reivew_p1': True,
       
        
        
        }
def pdfrw_get_fields():
    pdf_template="1_Pre6_form.pdf"
    pdf_output="new_form.pdf"
    template_pdf=pdfrw.PdfReader(pdf_template)
    
    ANNOT_KEY = '/Annots'
    ANNOT_FIELD_KEY = '/T'
    ANNOT_VAL_KEY = '/V'
    ANNOT_RECT_KEY = '/Rect'
    SUBTYPE_KEY = '/Subtype'
    WIDGET_SUBTYPE_KEY = '/Widget'
    
    for page in template_pdf.pages:
        annotations=page[ANNOT_KEY]
        for annotation in annotations:
            if annotation[SUBTYPE_KEY]== WIDGET_SUBTYPE_KEY:
                if annotation[ANNOT_FIELD_KEY]:
                    key= annotation[ANNOT_FIELD_KEY][1:-1]
                    
                    if key in employee_dict.keys():
                        print(key)
                        
                        if type(employee_dict[key])== bool:
                            if employee_dict[key]==True:
                                annotation.update(pdfrw.PdfDict(
                                   AS=pdfrw.PdfName('Yes')))
                        else:
                            
                            annotation.update(
                                pdfrw.PdfDict(V='{}'.format(employee_dict[key]))
                            )
                            annotation.update(pdfrw.PdfDict(AP=''))
    template_pdf.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))
    pdfrw.PdfWriter().write(pdf_output, template_pdf)

  
#get_fields()

def get_fields():
    #output=PyPDF2.PdfFileWriter("new_form.pdf")
    template=PyPDF2.PdfFileReader(open("1_Pre6_form.pdf", 'rb'))
    fields=template.getFields()
    print(fields)
    
get_fields()
#type_review_p1: 'appt', 'reappt'
#type_increase:"merit",'4th_year'
#######Notes
#/Ft= field type


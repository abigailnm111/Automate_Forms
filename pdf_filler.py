#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun May  2 21:33:27 2021

@author: panda
"""

import PyPDF2

employee_dict={
        
        'name_p1':' Regina George',
        'dept_p1': "Mathletes",
        'accrued_qtr':'10', #History record quarter count
        #'type_reivew_p1': True,
       'uid_p2':'12345678'
        
        
        }

radio_dict={
    'type_increase': '/4th_year',
    'type_review_p1':'/reappt'
    
    }


  


def get_fields():
    output=PyPDF2.PdfFileWriter()
    template=PyPDF2.PdfFileReader(open("1_Pre6_form.pdf", 'rb'))
    
    
    
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
    
  
    
    outputStream=open("NewPDF.pdf", "wb")
    output.write(outputStream)
   




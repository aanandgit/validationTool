'''
Created on Sep 17, 2016

@author: Hankock
'''
import os
import sys
import glob
import pprint
import openpyxl
import xmlMethods
import uatMethods
import Project_consts

#if (len(sys.argv) == 2):    
    
    #pwd = str(sys.argv[1])
pwd = os.getcwd()

if (len(pwd)!=0):    
    
    
    xml_files = []
    xlsx_files = [] 
    xml_data = {}
    xml_file_data = []
    
    for f in glob.glob(pwd + '\*.xml'):
        xml_files.append(f)
    
    for f in glob.glob(pwd + '\*.xlsx'):
        xlsx_files.append(f)
    #pprint.pprint(xml_files)   
    
    for i in range(0, len(xlsx_files)):
        if ((xlsx_files[i].find('UAT_')) != -1):
            Source_UAT = xlsx_files[i]
            
    
        
    # File Path of the generated file after parsing TS Reports and UAT
    ReportFilePath = pwd + Project_consts.Report_File_Name
    
    ## Parse source UAT to get all Test Cases
    wb_r = openpyxl.load_workbook(Source_UAT)
    uat_data = uatMethods.TestCasesInWorkbook(wb_r)
    #pprint.pprint(uat_data)
    
    ## Create ReportSummary File    
    wb_w = openpyxl.Workbook()
    wb_w.create_sheet(title = Project_consts.Sheet_name_1, index = 0)
    wb_w.create_sheet(title = Project_consts.Sheet_name_2, index = 1)
    wb_w.create_sheet(title = Project_consts.Sheet_name_3, index = 2)
    wb_w.remove_sheet(wb_w.get_sheet_by_name('Sheet'))
    wb_w.save(ReportFilePath)
        
    
    #Parse XML files     
    for i in range(0, len(xml_files)):
        xml_file_data.append(xmlMethods.DictFromXMLfile(xml_files[i]))
        xml_data = xmlMethods.mergeDictsIntoOne(xml_data,xml_file_data[i])
        
    #pprint.pprint(len(xml_data))
    #pprint.pprint(xml_data)
    
    # Generate Sheets of the Report File
    xmlMethods.CreateReportSheet(wb_r, wb_w, Project_consts.Sheet_name_2, xml_data,ReportFilePath)
    xmlMethods.CreateSummarySheet(wb_w, Project_consts.Sheet_name_1, uat_data, xml_data,ReportFilePath)
    xmlMethods.CreateTestsNotPerformed(wb_w,Project_consts.Sheet_name_3,uat_data, xml_data,ReportFilePath)
else:
    print('Invalid Inputs, Check Inputs and try again !!')
    print('UAT should be in .xlsx format only!!')
    print('Folder Path Name should be less than 100 characters, try placing the folder on Desktop')

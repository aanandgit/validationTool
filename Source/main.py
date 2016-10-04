'''
Created on Sep 17, 2016

@author: Hankock
'''
import openpyxl
import xmlMethods
import uatMethods
import pprint

xml_data = {}

wb_w = openpyxl.Workbook()

Sheet_name_1 = 'Summary'
Sheet_name_2 = 'Report'
Sheet_name_3 = 'NotAutomated'
# Report summary generated after parsing TS Reports and UAT
ReportFilePath = r'C:\Users\Hankock\Desktop\Python\Reports\Pump_400\Reports_CANOpen\ReportSummary.xlsx'
# Source UAT
Source_UAT = r'C:\Users\Hankock\Desktop\Python\Reports\Pump_400\Reports_CANOpen\UAT_PVXP_CANOpen.xlsx'

#wb_r = openpyxl.load_workbook(Source_UAT)
#pprint.pprint(names)


## Parse source UAT
wb_r = openpyxl.load_workbook(Source_UAT)
uat_data = uatMethods.TestCasesInWorkbook(wb_r)
#pprint.pprint(uat_data)
    

## Create ReportSummary File
wb_w.create_sheet(title = Sheet_name_1, index = 0)
wb_w.create_sheet(title = Sheet_name_2, index = 1)
wb_w.create_sheet(title = Sheet_name_3, index = 2)
wb_w.remove_sheet(wb_w.get_sheet_by_name('Sheet'))

wb_w.save(ReportFilePath)


## Parse xml_data files and create a dict of all the TestCases
file_path = r"C:\Users\Hankock\Desktop\Python\Reports\Pump_400\Reports_CANOpen\PXEP_TestBench_Report[6 12 15 PM][04 28 16].xml"
xml_data_1 = xmlMethods.DictFromXMLfile(file_path)

file_path = r"C:\Users\Hankock\Desktop\Python\Reports\Pump_400\Reports_CANOpen\PXEP_TestBench_Report[7 03 39 PM][04 28 16].xml"
xml_data_2 = xmlMethods.DictFromXMLfile(file_path)

file_path = r"C:\Users\Hankock\Desktop\Python\Reports\Pump_400\Reports_CANOpen\PXEP_TestBench_Report[7 53 32 PM][04 28 16].xml"
xml_data_3 = xmlMethods.DictFromXMLfile(file_path)

file_path = r"C:\Users\Hankock\Desktop\Python\Reports\Pump_400\Reports_CANOpen\PXEP_TestBench_Report[8 49 38 PM][04 28 16].xml"
xml_data_4 = xmlMethods.DictFromXMLfile(file_path)

file_path = r"C:\Users\Hankock\Desktop\Python\Reports\Pump_400\Reports_CANOpen\PXEP_TestBench_Report[9 25 52 PM][04 28 16].xml"
xml_data_5 = xmlMethods.DictFromXMLfile(file_path)

file_path = r"C:\Users\Hankock\Desktop\Python\Reports\Pump_400\Reports_CANOpen\PXEP_TestBench_Report[9 30 20 PM][04 28 16].xml"
xml_data_6 = xmlMethods.DictFromXMLfile(file_path)

xml_data = xmlMethods.mergeDictsIntoOne(xml_data,xml_data_1)
xml_data = xmlMethods.mergeDictsIntoOne(xml_data,xml_data_2)
xml_data = xmlMethods.mergeDictsIntoOne(xml_data,xml_data_3)
xml_data = xmlMethods.mergeDictsIntoOne(xml_data,xml_data_4)
xml_data = xmlMethods.mergeDictsIntoOne(xml_data,xml_data_5)
xml_data = xmlMethods.mergeDictsIntoOne(xml_data,xml_data_6)

"""
xml_data = xml_data_1.copy() 
xml_data.update(xml_data_2)
"""
#pprint.pprint(len(xml_data))
#pprint.pprint(xml_data)

xmlMethods.CreateReportSheet(wb_r, wb_w, Sheet_name_2, xml_data,ReportFilePath)

xmlMethods.CreateSummarySheet(wb_w, Sheet_name_1, uat_data, xml_data,ReportFilePath)

xmlMethods.CreateTestsNotPerformed(wb_w,Sheet_name_3,uat_data, xml_data,ReportFilePath)

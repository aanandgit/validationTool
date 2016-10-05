'''
Created on Sep 17, 2016

@author: Hankock
'''
import xml.etree.ElementTree as ET
import pprint
import uatMethods
from openpyxl.styles import Alignment, Color

#===============================================================================
# This function creates the Summary sheet. This sheet gives a summary of the 
# UAT and the Test Stand Reports.
#===============================================================================
def CreateSummarySheet(wb_w, Sheet_name, uat_data, xml_data,ReportFilePath):
    Summary_sheet = wb_w.get_sheet_by_name(Sheet_name)
    
    #===========================================================================
    # Data related to the source UAT
    #===========================================================================
    
    ur = 5
    uc = 5
    
    NoOfTCsInUAT = len(uat_data)
    NoOfTCsInTS = len(xml_data)
    NoOfTCsNotAutomated = (len(uat_data) - len(xml_data))
    
    Summary_sheet.cell(row = ur, column = uc).value = 'Total No of TCs in UAT'
    Summary_sheet.cell(row = ur, column = uc+1).value = NoOfTCsInUAT
    Summary_sheet.cell(row = ur+1, column = uc).value = 'Total No of TCs in TestStand Reports'
    Summary_sheet.cell(row = ur+1, column = uc+1).value = NoOfTCsInTS
    Summary_sheet.cell(row = ur+2, column = uc).value = 'TCs Not Automted'
    Summary_sheet.cell(row = ur+2, column = uc+1).value = NoOfTCsNotAutomated
    
    
    #===========================================================================
    # Data related to the TestStand Reports
    #===========================================================================
    
    xr = 5
    xc = 2
    
    l_TCsPassFailSkip = TCsPassFailSkip(xml_data)
    
    nNoOfTCsPassed = len(l_TCsPassFailSkip[0])
    nNoOfTCsFailed = len(l_TCsPassFailSkip[1])
    nNoOfTCsSkiped = len(l_TCsPassFailSkip[2])
    nTotalNoOfTcs  = nNoOfTCsPassed + nNoOfTCsFailed + nNoOfTCsSkiped
     
    Summary_sheet.cell(row = xr, column = xc).value = 'TCs Passed'
    Summary_sheet.cell(row = xr, column = xc+1).value = nNoOfTCsPassed
    Summary_sheet.cell(row = xr+1, column = xc).value = 'TCs Failed'
    Summary_sheet.cell(row = xr+1, column = xc+1).value = nNoOfTCsFailed
    Summary_sheet.cell(row = xr+2, column = xc).value = 'TCs Skipped'
    Summary_sheet.cell(row = xr+2, column = xc+1).value = nNoOfTCsSkiped
    Summary_sheet.cell(row = xr+3, column = xc).value = 'Total No Of TestCases'
    Summary_sheet.cell(row = xr+3, column = xc+1).value = nTotalNoOfTcs
    
    uatMethods.setColumnWidth(Summary_sheet)
    '''if nNoOfTCsFailed !=0:
        Summary_sheet.cell(row = xr+1, column = xc+1) = color()
        '''
    wb_w.save(ReportFilePath)


#===============================================================================
# This function creates Report Sheet. This sheet lists all the Test Cases in
# the TestStand Report Files.
#===============================================================================
def CreateReportSheet(wb_r, wb_w, Sheet_name, data, ReportFilePath):
      
    Report_sheet = wb_w.get_sheet_by_name(Sheet_name)
    
    names = uatMethods.SheetNamesbyTCnames(wb_r)
    
    wr_column_name = True
    prev_sheet_no = 1 # should be init at runtime
    curr_sheet_no = 0
    r = 3
    c = 2
    
    for i,j in sorted(data.items()): #Does it sort by name..???
        tc_name = str(i)
        sh_name_key, middle, last = tc_name.partition('_')
        sh_name_value = names[sh_name_key]
        
        #print(sh_name_key + ' ' +  names[sh_name_key])
                        
        curr_sheet_no = int(tc_name[2:5])   
        if (curr_sheet_no != prev_sheet_no):
            c +=3
            r = 3
            wr_column_name = True
            prev_sheet_no =  curr_sheet_no

        if wr_column_name == True:
            Report_sheet.cell(row = r-1, column = c).value = str(sh_name_value)
            Report_sheet.merge_cells(start_row=r-1, start_column=c, end_row=r-1, end_column=c+1)
            Report_sheet.cell(row=r-1, column=c).alignment = Alignment(horizontal = 'center')
            wr_column_name = False
        Report_sheet.cell(row = r, column=c).alignment = Alignment(horizontal = 'center')
        Report_sheet.cell(row = r, column = c).value = i
        Report_sheet.cell(row = r, column = c+1).value = j
        r +=1
    
    uatMethods.setColumnWidth(Report_sheet) 
    
    wb_w.save(ReportFilePath)    

#=======================================================================
# This function creates separate dict of Passed, Failed and Skipped 
# test cases from the TestStand Report files
#=======================================================================

def CreateTestsNotPerformed(wb_w, Sheet_name, uat_data, xml_data, ReportFilePath):
    
    Sheet = wb_w.get_sheet_by_name(Sheet_name)
    
    xml_list = xml_data.keys()
    not_list = uat_data.keys()
    len_not_list = len(not_list)
    len_xml_list = len(xml_list)
    
    
    for i in range(0, len_xml_list):
        for j in range(0, len_not_list):
            if xml_list[i] == not_list[j]:
                del not_list[j]
                len_not_list -=1
                break

    Sheet.column_dimensions['B'].width  = 22                  
    Sheet.column_dimensions['D'].width  = 22
    Sheet.column_dimensions['F'].width  = 22
    
    Sheet.cell(row = 2, column = 2).value = "TCs not automated"
    Sheet.cell(row = 2, column = 4).value = "TCs in TestStand Reports"
    Sheet.cell(row = 2, column = 6).value = "TCs in the UAT"
    
    for num in range(0, len(not_list)):
        Sheet.cell(row = num+2, column = 2).alignment = Alignment(horizontal = 'center')
        Sheet.cell(row = num+3, column = 2).value = not_list[num]
    for num in range(0, len(xml_data)):
        Sheet.cell(row = num+2, column = 4).alignment = Alignment(horizontal = 'center')
        Sheet.cell(row = num+3, column = 4).value = xml_data.keys()[num]
    for num in range(0, len(uat_data)):
        Sheet.cell(row = num+2, column = 6).alignment = Alignment(horizontal = 'center')
        Sheet.cell(row = num+3, column = 6).value = uat_data.keys()[num] 
        
    #print(len(tc_list))
    #pprint.pprint(tc_list)
    
    uatMethods.setColumnWidth(Sheet)
    wb_w.save(ReportFilePath)



#===============================================================================
# Take dicts of infividual xml files and merhe them into 1 mega dict
#===============================================================================
def mergeDictsIntoOne(master_dict, new_dict):
    key_found = False
    if len(master_dict) == 0:
        master_dict = new_dict
    else:
        for new_key, new_value in new_dict.items():
            for master_key, master_value in master_dict.items():
                if new_key == master_key:
                    if master_value == 'Failed':
                        # It has been marked Failed in master dict,
                        # Update it only if the status has changed to Passed
                        if new_value == 'Passed':
                            master_value = new_value
                            key_found = True
                            break
                    elif master_value == 'Skipped':
                        # It has been marked Skipped in the master dict, 
                        # Update it only if the new status is either Passed or Failed
                        if new_value != 'Skipped':
                            master_value = new_value
                            key_found = True
                            break
 
            if (key_found == False):
                master_dict.setdefault(new_key, new_value)
            elif(key_found == True):
                master_dict[master_key] = master_value
                    
    #print(len(master_dict))                    
    return master_dict                
   
                
                
    
#===============================================================================
# This function creates separate lists of Pass Fail and Skipped TCs from 
# TS reports
#===============================================================================
def TCsPassFailSkip(xml_data):
    passed = {}
    failed = {}
    skipped = {}
    unknown = {}
    for i,j in sorted(xml_data.items()):
        value = str(j)
        if value == 'Passed':
            passed.setdefault(i,j)            
        elif value == 'Failed':
            failed.setdefault(i,j)
        elif value == 'Skipped':
            skipped.setdefault(i,j)
        else:
            unknown.setdefault(i,j)
    return passed,failed,skipped,unknown       


#===========================================================================
# This function parse a TestStand Report file and generates a dict of all 
# the test cases and their status.
#===========================================================================
def DictFromXMLfile(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()
    
    tc_flag = False  ## Found a TC Name
    status_flag = False  ## Found a Status Value
    store_tc = 'NA'
    store_result = 'NA'
    result = 'NA'
    tc = 'NA'
    result_list = []
    tc_list = []
    report = {}
    
    for i in root.iter('Prop'):
        if i.attrib.get('Name', '0') == 'StepName':
            tc = i.find('Value')
            tc = str(tc.text)
            if tc.startswith('TC'):
                tc_flag = True  # found a TC name, should look for Status now
                store_tc = tc
          
        if i.attrib.get('Name', '0') == 'Status':
            result = i.find('Value')
            result = str(result.text)
            if result == 'Passed' or result == 'Failed' or result == 'Skipped':
                status_flag = True
                store_result = result
    
                
        if ((tc_flag == True) and (status_flag == True)):
            tc_flag = False
            status_flag = False
            tc_list.append(store_tc)
            result_list.append(store_result)
        
    '''
    print (len(tc_list))
    print(len(result_list))
    pprint.pprint(tc_list)
    pprint.pprint(result_list)
    
    '''
    for i in range(0, len(tc_list)):
        report.setdefault(tc_list[i], result_list[i])
        
    # pprint.pprint(report)
    return report

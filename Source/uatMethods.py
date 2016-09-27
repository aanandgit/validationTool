'''
Created on Sep 10, 2016

@author: Hankock
'''
from openpyxl.styles import Font, Alignment

def CreateSummarySheetFromData(wb_read, wb_write, Data, file_name):
   
    Sheet_write = wb_write.get_sheet_by_name('Summary')
    
    Count_TCs = TestCasesCountInAllSheets(wb_read)   
    
    ft = Font(bold=True)
    al = Alignment(horizontal='center')
    
    Sheet_write.cell(row=2, column=2).value = 'Sheet Name'
    Sheet_write.cell(row=2, column=3).value = 'No of Testcases'
    
    
    Sheet_write.cell(row=2, column=2).font = ft
    Sheet_write.cell(row=2, column=3).font = ft
    
    Sheet_write.cell(row=2, column=2).alignment = al
    Sheet_write.cell(row=2, column=3).alignment = al
    
    Sheet_write.column_dimensions['B'].width = 35
    Sheet_write.column_dimensions['C'].width = 20
    
    r = 3
    c = 2
    
    # for i in sorted(Count_TCs.items()):
    for key, value in sorted(Count_TCs.items()):
        Sheet_write.cell(row=r, column=c).value = key
        Sheet_write.cell(row=r, column=c + 1).value = value
        Sheet_write.cell(row=r, column=c + 1).alignment = al
        r += 1
        
    
    wb_write.save(file_name)
     
def CreateReportSheetFromData(wb_read, wb_write, Data, file_name):
    
    Sheet_write = wb_write.get_sheet_by_name('Report')
    
    List_of_sheets = wb_read.get_sheet_names()
        
    c = 2  # column Iterator
    k = 3  # row Iterator
    s = 1  # Sheet Iterator
    n = 0  # no of rows written
   
       
    No_of_TCs = TestCasesCountInSheet(wb_read, List_of_sheets[s])
    
    for i, j in sorted(Data.items()):  # sorts the dict before iterating
        Sheet_write.cell(row=k, column=c).value = i
        Sheet_write.cell(row=k, column=c + 1).value = j
        
        n = n + 1
        if n == No_of_TCs:
            n = 0
            c = c + 2
            k = 3
            s = s + 1 
            No_of_TCs = TestCasesCountInSheet(wb_read, List_of_sheets[s])
        else:
            k = k + 1            
    # print(Sheet_write.cell(row = i, column = 2).value + ' ' + Sheet_write.cell(row = i, column = 3).value)

    wb_write.save(file_name)
    
def TestCasesInWorkbook(Workbook):
    List_of_sheets = Workbook.get_sheet_names()
    Num_of_sheets = len(List_of_sheets)
    
    Sheet_data = {}
    Data_TCs = {}
    
    for i in range(1, Num_of_sheets):
        Sheet_data = TestCasesInSheet(Workbook, List_of_sheets[i])
        for j in range(0, len(Sheet_data)):
            Data_TCs.setdefault(Sheet_data.keys()[j], Sheet_data.values()[j])
        
    return Data_TCs
          
def TestCasesInSheet(Workbook, Sheet):
    Index = GetDataIndex(Workbook, Sheet, 'TestCase Name')
    column_num = Index.values()[0]
    Sheet_data = {}
    Sheet = Workbook.get_sheet_by_name(Sheet)    
    for i in range(1, (Sheet.max_row+1)):
        row_value = str(Sheet.cell(row=i, column=column_num).value)
        if (row_value.startswith('TC') == True):
            Sheet_data.setdefault(row_value, 'False') 
                                    
    return Sheet_data

def TestCasesCountInAllSheets(Workbook):
    List_of_sheets = Workbook.get_sheet_names()
    Num_of_sheets = len(List_of_sheets)
    Count_TCs = {}
    
    for i in range(1, Num_of_sheets):
        Num_of_TCs = TestCasesCountInSheet(Workbook, List_of_sheets[i])
        Count_TCs.setdefault(List_of_sheets[i], Num_of_TCs)

    return Count_TCs

def TestCasesCountInSheet(Workbook, Sheet):
             
    No_of_TCs = 0
    # Get Index of Test Case Name  in the current sheet
    # Get column no of TestCase number
    Index = GetDataIndex(Workbook, Sheet, 'TestCase Name')
    column_num = Index.values()[0]
       
    Sheet = Workbook.get_sheet_by_name(Sheet)
   
    for j in range(1, Sheet.max_row):
        row_value = str(Sheet.cell(row=j, column=column_num).value)
        if (row_value.startswith('TC') == True):
            No_of_TCs = No_of_TCs + 1
       
    return No_of_TCs

def GetDataIndex(Workbook, Sheet, data):
    Index = {}
    Sheet = Workbook.get_sheet_by_name(Sheet)
    for i in range(1, Sheet.max_row):
        for j in range(1, Sheet.max_column):
            column_value = Sheet.cell(row=i, column=j).value
            # column_value and data should have same case ..???
            # print(column_value)
            if (str(column_value) == str(data)):  # does it need to be in str..???
                
                break
        break   
    Index.setdefault(i, j)  # should be integer value only
       
    return Index

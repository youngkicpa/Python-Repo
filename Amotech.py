import win32com.client as win
import psutil
from GLforAmotech import *
from MyPPrint import *
from Voucher import *
from Vouchers import *

def killExcel():
    for proc in psutil.process_iter():
        # check whether the process name matches
            if proc.name() == "EXCEL.EXE":
                proc.kill()

def LoadExcel(filename, sheetname):
    xl = win.gencache.EnsureDispatch("Excel.Application")
    xl.Visible = False

    wb = xl.Workbooks.Open(filename)
    ws = wb.Worksheets(sheetname)

    data = ws.UsedRange.Value

    wb.Save()
    wb.Close()
    xl.Quit()
    
    return data

def SaveExcel(filename, data):
    xl = win.gencache.EnsureDispatch("Excel.Application")
    xl.Visible = False

    wb = xl.Workbooks.Add()
    ws = wb.Worksheets(1)

    ws.Range(xl.Cells(1, 1), xl.Cells(len(data), len(data[0]))).Value = tuple(data)

    wb.SaveAs(filename)
    wb.Close()
    xl.Quit()
    
    return data

if __name__ == "__main__":
    filename = "C:\\DataTest\\분개장_FY23_아모텍.xlsx"
    sheetname = "2023"
    targetFileName = "C:\\Users\\young\\Downloads\\result.xlsx"
    
    data = LoadExcel(filename, sheetname)
    killExcel()
    
    #glforAmotech = Amotech(data)

    #glforAmotech.getTrialBalance()

    #trialBalance = glforAmotech.changeTrialBalanceList()
    #SaveExcel(targetFileName, sorted(trialBalance))
    #glforAmotech.printTrialBalance()
    
    vouchers = Vouchers()
    vouchers.getVouchers(data)

    vouchers.getVouchersAmounts()
    #result = vouchers.getFiltered(vouchers.testVoucherMinusSales)   
    #vouchers.salesTransactions()

    #filtered = minusSalesTransactions(vouchers)
    
    #filteredList = []
    
    #for voucher in filtered:
    #     filteredList.extend(voucher.ToList())

    #SaveExcel(targetFileName, filteredList)
    #del data
    #testVoucherAmount(vouchers)

    #result = vouchers.getVouchersAmounts()
    #SaveExcel(targetFileName, result)
    print("Hello")


    

    
    
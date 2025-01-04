# %%
import win32com.client as win
import psutil
import os
from SamjinVoucher import *
from SamjinVouchers import *
from MyPPrint import *


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

    listOfFile = os.listdir(os.path.dirname(filename))
    wb = None
    ws = None

    if os.path.basename(filename) in listOfFile:
        wb = xl.Workbooks.Open(filename)
        ws = wb.Worksheets.Add()
    else:
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets(1)

    ws.Range(xl.Cells(1, 1), xl.Cells(len(data), len(data[0]))).Value = tuple(data)

    wb.SaveAs(filename)
    wb.Close()
    xl.Quit()
    killExcel()
    return data

def GetExcelData():
    filename = "C:\\DataTest\\삼진엘앤디_분개장_상반기_yskim.xlsx"
    sheetname = "Sheet1"  
    data = LoadExcel(filename, sheetname)
    killExcel()
    return data

def GetMinusSalesData(data):
    filltered = SJVouchersData.getFiltered(SJVouchersData.testVoucherMinusSales)
    result = []
    result.append(data[0])
    for i, x in enumerate(data):
        if i == 1:
            continue
        else:
            if x[12] in filltered:
                result.append(x) 
    SaveExcel(targetFileName, result)

def GetSalesForReason(voucher):
    if 52000102 in voucher.debit["codes"]:
        for i, y in enumerate(voucher.debit["codes"]):
            if y == 52000102 and voucher.debit["amounts"][i-1] > 0:
                for x in voucher.credit["codes"]:
                    if 51000101 <= int(x) <= 51000502:
                        return True 
    return False

def GetFilteredData(data, SJVouchersData):
    filltered = SJVouchersData.getFiltered(GetSalesForReason)
    result = []
    result.append(data[0])
    for i, x in enumerate(data):
        if i == 1:
            continue
        else:
            if x[12] in filltered:
                result.append(x) 
    SaveExcel(targetFileName, result)

# %%
if __name__ == "__main__":    
    targetFileName = "C:\\Users\\young\\Downloads\\SamjinResult.xlsx"
    data = GetExcelData()

    SJVouchersData = SamjinVouchers()
    SJVouchersData.getVouchers(data)

    # 합계잔액시산표를 분개장으로부터 추출하는 것  
    # -----------------------------
    #SJVouchersData.getTrialBalance()    
    #SaveExcel(targetFileName, SJVouchersData.getTrialBalanceTuple())

    GetFilteredData(data, SJVouchersData)
    #GetMinusSalesData(data)
    #result = SJVouchersData.minusSalesTransactions()
    #SaveExcel(targetFileName, result)
    print("Done")  



# %%

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

def GetExcelData():
    filename = "C:\\DataTest\\아모텍_분개장 (2024_06).xlsx"
    sheetname = "Sheet1"  
    data = LoadExcel(filename, sheetname)
    killExcel()
    return data

if __name__ == "__main__":    
    targetFileName = "C:\\Users\\young\\Downloads\\result.xlsx"
    data = GetExcelData()
    vouchers = Vouchers()
    vouchers.getVouchers(data)

    
    trialNo = 0
    while  1:
        print("다음 중 원하는 작업을 선택하시요")
        print("   1. 합계시산표를 만들기")
        print("   2. 매출이 (-)인 전표를 추출하기")
        print("   3. 매출전표들의 차대변 합계를 구하기")
        print("   4. 차대변 합계가 다른 전표 확인하기")
        print("   5. 전표금액의 범위별 숫자확인하기")
        print("   9. 종료하기")
        print("   \nAttributeError: module 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9' has no attribute 'CLSIDToClassMap'")
        print("   위의 에러메시지가 발생을 하면, 해결하는 방법은 C:\\Users\\young\\AppData\\Local\\Temp\\gen_py 폴더의 내용을 모두 삭제한다.")
        print("   그래도 안되면, pip uninstall pywin32 그리고 pip install pywin32")
        selection = input()
        trialNo += 1
        match selection:
            case '1':
                print(f"\nHello {trialNo}")
                glforAmotech = Amotech(data)
                glforAmotech.getTrialBalance()
                trialBalance = glforAmotech.changeTrialBalanceList()
                SaveExcel(targetFileName, sorted(trialBalance))
                glforAmotech.printTrialBalance()
            case '2':
                print(f"\nHello {trialNo}")
                filtered = vouchers.minusSalesTransactions()    
                filteredList = []    
                for voucher in filtered:
                     filteredList.extend(voucher.ToList())
                SaveExcel(targetFileName, filteredList)
            case '3':
                print(f"\nHello {trialNo}")
                vouchers.salesTransactions()
            case '4':
                print(f"\nHello {trialNo}")
                vouchers.testVoucherAmount()
            case '5':
                print(f"\nHello {trialNo}")           
                result = vouchers.getVouchersAmounts()
                SaveExcel(targetFileName, result)
            case '9':
                break 
            case _ : 
                continue 

        

    

    
    
import win32com.client as win
import psutil
from GLforAmotech import *
from MyPPrint import *
from Voucher import *

def killExcel():
    for proc in psutil.process_iter():
        # check whether the process name matches
            if proc.name() == "EXCEL.EXE":
                proc.kill()


def LoadExcle(filename, sheetname):
    xl = win.gencache.EnsureDispatch("Excel.Application")
    xl.Visible = False

    wb = xl.Workbooks.Open(filename)
    ws = wb.Worksheets(sheetname)

    data = ws.UsedRange.Value

    wb.Save()
    wb.Close()
    xl.Quit()
    
    return data

def testforTitles(row):
        if row[0] == "회계일":
            return True
        else:
            return False
        
def getVoucher(data):
    Vouchers = []
    previousNo = ""
    start = False
    currentVoucher = None
    for row in data:
        if testforTitles(row):
            start = True
            currentVoucher = Voucher()
            continue
        if not start:
            continue
        else:
            if row[1][0:13] == previousNo:
                    currentVoucher.Add(row)
            else:
                    Vouchers.append(currentVoucher)
                    currentVoucher = Voucher()
                    currentVoucher.Add(row)
                    previousNo = row[1][0:13]

    return Vouchers
                     
def testVoucherAmount(vouchers):
    for voucher in vouchers:
         if not voucher.TestAmounts():
              print(f"{voucher.no}\t {voucher.debitSum}\t{voucher.creditSum}\t{voucher.TestAmounts()}")



if __name__ == "__main__":
    filename = "C:\\DataTest\\분개장_FY23_아모텍.xlsx"
    sheetname = "2023"
    data = LoadExcle(filename, sheetname)
    killExcel()
    glforAmotech = Amotech(data)

    #glforAmotech.getTrialBalance()
    #glforAmotech.printTrialBalance()
    vouchers = getVoucher(data)
    
    del data
    testVoucherAmount(vouchers)

    print("Hello")


    

    
    
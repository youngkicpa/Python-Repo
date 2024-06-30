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

def testVoucherSales(voucher):
    if len(voucher.credit["codes"]) == 0:
         return False
    for x in voucher.credit["codes"]:
        if 4110000 <= int(x) <= 4200000:
            return True 
    return False

def testVoucherMinusSales(voucher):
    if len(voucher.credit["codes"]) == 0:
         return False
    for index, x in enumerate(voucher.credit["codes"]):
        if 4110000 <= int(x) <= 4200000 and voucher.credit["amounts"][index] < 0:
            return True 
    return False

def salesTransactions(vouchers):
    debitInfo = {}
    creditInfo = {}
    for voucher in vouchers:
        if testVoucherSales(voucher):
            for index, d in enumerate(voucher.debit["accounts"]):
                if d in debitInfo:
                    debitInfo[d] += voucher.debit["amounts"][index]
                else:
                    debitInfo[d] = voucher.debit["amounts"][index]

            for index, c in enumerate(voucher.credit["accounts"]):
                if c in creditInfo:
                    creditInfo[c] += voucher.credit["amounts"][index]
                else:
                    creditInfo[c] = voucher.credit["amounts"][index]

    print("차변")
    for key, value in debitInfo.items():
         print(f"{key}\t{value}")
    print("대변")
    for key, value in creditInfo.items():
         print(f"{key}\t{value}")

def minusSalesTransactions(vouchers):
    debitInfo = {}
    creditInfo = {}
    count = 0
    filtered = []
    for voucher in vouchers:
        if testVoucherMinusSales(voucher):
            count += 1
            filtered.append(voucher)
            for index, d in enumerate(voucher.debit["accounts"]):
                if d in debitInfo:
                    debitInfo[d] += voucher.debit["amounts"][index]
                else:
                    debitInfo[d] = voucher.debit["amounts"][index]

            for index, c in enumerate(voucher.credit["accounts"]):
                if c in creditInfo:
                    creditInfo[c] += voucher.credit["amounts"][index]
                else:
                    creditInfo[c] = voucher.credit["amounts"][index]

    print(f"{count}개의 전표가 있습니다.")
    print("차변")
    for key, value in debitInfo.items():
         print(f"{key}\t{value}")
    print("대변")
    for key, value in creditInfo.items():
         print(f"{key}\t{value}")
    return filtered

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
    data = LoadExcel(filename, sheetname)
    killExcel()
    #glforAmotech = Amotech(data)

    #glforAmotech.getTrialBalance()
    #glforAmotech.printTrialBalance()
    vouchers = getVoucher(data)
    
    filtered = minusSalesTransactions(vouchers)
    filteredList = []
    for voucher in filtered:
         filteredList.extend(voucher.ToList())

    SaveExcel("C:\\Users\\young\\Downloads\\result.xlsx", filteredList)
    del data
    #testVoucherAmount(vouchers)

    print("Hello")


    

    
    
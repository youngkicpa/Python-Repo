# %%
import win32com.client as win
import psutil
import os
from SamjinVoucher2ndHalf import *
from SamjinVouchers2ndHalf import *
#from MyPPrint import *


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
    filename = r"C:\Users\young\Downloads\계정별원장_삼진엘앤디_하반기.xlsx"
    sheetname = "계정별원장"  
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
    targetFileName = "C:\\Users\\young\\Downloads\\SamjinResult2nd.xlsx"
    data = GetExcelData()

    SJVouchersData = SamjinVouchers()
    SJVouchersData.getVouchers(data)

    #GetFilteredData(data, SJVouchersData)
    #GetMinusSalesData(data)
    #result = SJVouchersData.minusSalesTransactions()
    #SaveExcel(targetFileName, result)
    #print("Done")  

    trialNo = 0
    while  1:
        print("다음 중 원하는 작업을 선택하시요")
        print("   1. 합계시산표를 만들기")
        print("   2. 매출이 (-)인 전표를 추출하기")
        print("   3. 매출전표들의 차대변 합계를 구하기")
        print("   4. 차대변 합계가 다른 전표 확인하기")
        print("   5. 전표금액의 범위별 숫자확인하기")
        print("   6. 분개장 총차대변 금액확인하기")
        print("   9. 종료하기")
        print("   \nAttributeError: module 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9' has no attribute 'CLSIDToClassMap'")
        print("   위의 에러메시지가 발생을 하면, 해결하는 방법은 C:\\Users\\young\\AppData\\Local\\Temp\\gen_py 폴더의 내용을 모두 삭제한다.")
        print("   그래도 안되면, pip uninstall pywin32 그리고 pip install pywin32")
        selection = input("선택:  ")
        trialNo += 1
        match selection:
            case '1':
                print(f"\nHello {trialNo}")
                SJVouchersData.getTrialBalance()    
                SaveExcel(targetFileName, SJVouchersData.getTrialBalanceTuple())
            case '2':
                print(f"\nHello {trialNo}")
                test = False
                result = []
                debit = {}
                credit = {}
                for voucher in SJVouchersData.vouchers:
                    for i, code in enumerate(voucher.credit["codes"]):
                        if (51000100 < int(code) < 51000600) and voucher.credit["amounts"][i] < 0:
                            test = True
                            break
                    if test == False:
                        for code in voucher.debit["codes"]:
                            if (51000100 < int(code) < 51000600) and voucher.debit["amounts"][i] > 0:
                                test = True
                                break
                    if test == True:
                        for i, code in enumerate(voucher.credit["codes"]):
                            if code in credit.keys():
                                credit[code][1] += voucher.credit["amounts"][i]
                            else:
                                credit[code] = []
                                credit[code].append(voucher.credit["accounts"][i])
                                credit[code].append(voucher.credit["amounts"][i])
                        for i, code in enumerate(voucher.debit["codes"]):
                            if code in debit.keys():
                                debit[code][1] += voucher.debit["amounts"][i]
                            else:
                                debit[code] = []
                                debit[code].append(voucher.debit["accounts"][i])
                                debit[code].append(voucher.debit["amounts"][i])
                    test = False
                
                for key, value in debit.items():
                    result.append(("차변", key, value[0], value[1]))
                for key, value in credit.items():
                    result.append(("대변", key, value[0], value[1]))
                SaveExcel(targetFileName, result)
                print("Done")                
            case '3':
                print(f"\nHello {trialNo}")
                test = False
                result = []
                debit = {}
                credit = {}
                for voucher in SJVouchersData.vouchers:
                    for code in voucher.credit["codes"]:
                        if 51000100 < int(code) < 51000600:
                            test = True
                            break
                    if test == False:
                        for code in voucher.debit["codes"]:
                            if 51000100 < int(code) < 51000600:
                                test = True
                                break
                    if test == True:
                        for i, code in enumerate(voucher.credit["codes"]):
                            if code in credit.keys():
                                credit[code][1] += voucher.credit["amounts"][i]
                            else:
                                credit[code] = []
                                credit[code].append(voucher.credit["accounts"][i])
                                credit[code].append(voucher.credit["amounts"][i])
                        for i, code in enumerate(voucher.debit["codes"]):
                            if code in debit.keys():
                                debit[code][1] += voucher.debit["amounts"][i]
                            else:
                                debit[code] = []
                                debit[code].append(voucher.debit["accounts"][i])
                                debit[code].append(voucher.debit["amounts"][i])
                    test = False
                
                for key, value in debit.items():
                    result.append(("차변", key, value[0], value[1]))
                for key, value in credit.items():
                    result.append(("대변", key, value[0], value[1]))
                SaveExcel(targetFileName, result)
                print("Done")
            case '4':
                print(f"\nHello {trialNo}")
                result = 0
                errors = []
                for voucher in SJVouchersData.vouchers:
                    if voucher.creditSum != voucher.debitSum:
                        result += 1
                        errors.append(voucher.no)
                if result == 0:
                    print("전표의 합계검증이 완료되었습니다. 차대변이 일치하지 않는 전표가 없습니다.")
                else:
                    print(f"{result}개의 전표 차대변이 일치하지 않습니다.")
                    print(errors)
            case '5':
                print(f"\nHello {trialNo}")
                SJVouchersData.getVouchersAmounts()
            case '6':
                print(f"\nHello {trialNo}")
                debitSum = 0
                creditSum = 0
                for voucher in SJVouchersData.vouchers:
                    debitSum += voucher.debitSum
                    creditSum += voucher.creditSum
                print(f"차변합계: {debitSum} \t대변합계: {creditSum}")
            case '9':
                break 
            case _ : 
                continue 


# %%

# %%
import os
import tabula
import numpy as np
import psutil
import win32com.client as win

def killExcel():
    for proc in psutil.process_iter():
        # check whether the process name matches
            if proc.name() == "EXCEL.EXE":
                proc.kill()

def loadExcel(filename, sheetname):
    xl = win.gencache.EnsureDispatch("Excel.Application")
    xl.Visible = False

    wb = xl.Workbooks.Open(filename)
    ws = wb.Worksheets(sheetname)

    data = ws.UsedRange.Value

    wb.Save()
    wb.Close()
    xl.Quit()
    
    return data

def saveExcel(filename, data):
    xl = win.gencache.EnsureDispatch("Excel.Application")
    xl.Visible = False

    wb = xl.Workbooks.Add()
    ws = wb.Worksheets(1)

    ws.Range(xl.Cells(1, 1), xl.Cells(len(data), len(data[0]))).Value = tuple(data)

    wb.SaveAs(filename)
    wb.Close()
    xl.Quit()
    
    return data

def getExcelData():
    filename = "C:\\DataTest\\아모텍_분개장 (2024_06).xlsx"
    sheetname = "Sheet1"  
    data = loadExcel(filename, sheetname)
    killExcel()
    return data

listOfCategory = [
    "1. 조회기준일 현재 조회대상회사의 당 은행에 대한 금융상품의 내용은 다음과 같습니다.",
    "2. 조회기준일 현재 조회대상회사에 대한 당 은행의 대출거래의 내용은 다음과 같습니다.",
    "3. 조회기준일 현재 조회대상회사에 대한 당 은행의 지급보증 및 기타 약정사항의 내용은 다음과 같습니다.",
    "4. 조회기준일 현재 조회대상회사의 미결제파생상품계약 등(선물환, 스왑, 옵션, 기타 이와 유사한 계약 포함)의 내용은 다음과 같습니다.",
    "5. 조회기준일 현재 조회대상회사가 타 법인(개인)을 위하여 당행 앞으로 제공한 담보 및 연대보증의 내용은 다음과 같습니다.",
    "6. 당 은행이 2024년 01월 01일 부터 2023년 12월 31일까지 조회대상에 교부한 전자어음, 어음수표의 용지는 다음과 같습니다.",
    "7. 당 은행이 조회대상회사에게 교부한 전자어음과 어음·수표 중 조회기준일 현재 미발행되거나 미결제된 전자어음과 미회수된 어음·수표의 내역은 다음과 같습니다.",
    "8. 당 은행이 조회대상회사로부터 조회기준일 현재 담보, 견질 목적으로 보관하고 있는 어음이나 수표의 내역은 다음과 같습니다.",
    "9. 조회기준일 현재 조회대상회사의 당 은행에 대한 대출금 등 모든 신용공여와 관련하여 조회 대상회사의 자산 등이 당 은행에 담보와 보증 등으로 제공된 내역과 제 3자로부터 제공받은 담보와 보증 등의 내역은 다음과 같으며, 이는 참고 목적으로 제공되므로 그 정확성을 보증할 수는 없습니다.",
    "10. 2023년 01월 01일 부터 2023년 12월 31일 까지 조회대상회사가 거래한 당좌거래명세는 다음과 같습니다."
]

def getDataFromPDF(filePath):
    if os.path.exists(filePath):
        print("File exists, proceeding with PDF processing.")
        # PDF 파일 읽기
        dfs = tabula.read_pdf(filePath, stream=True, pages='all')
        print(f"Number of dataframes: {len(dfs)}")
        return dfs
    else:
        print(f"File not found: {filePath}")

def printPDFData(dfs):
    index = 1
    if len(dfs) > 0:
        for x in dfs:
            print(f"{index}번째")
            if x.empty:
                print(tuple(x.columns.to_list()))
            else:
                print(tuple(x.columns.to_list()))
                for y in x.values:
                    print(tuple(y))
            index += 1

def processCol(source, data):
    column = 0
    for ele in source:
        if isinstance(ele, str) and ele.startswith("Unnamed:"):
            data[column] = None
        else:
            data[column] = ele
        column += 1

def testPDFData(dfs):
    index= 1
    if len(dfs) > 0:
        for x in dfs:
            print(index)
            if x.empty:
                print("Empty DataFrame")
                print(f"컬럼: {type(x.columns)}")
            else:
                print(f"컬럼: {type(x.columns)}")
                print(f"Values: {type(x.values)}")
                print(f"Shape: {type(x.values.shape)}")
                print(f"Values[0]: {type(x.values[0])}")
            index += 1

def testSkipData(dfs):
    if len(dfs) > 0:
        if dfs.empty:
            return False
        for x in tuple(dfs.values):
            print(x)
            if isinstance(x, str) or isinstance(x, float):
                return False
            elif isinstance(x[0], float):
                return False
            elif x[0].endswith("전자어음"):
                return True
            else:
                continue
    else:
        return False
    
    return False

import numpy as np

def processPDFData(dfs, data):
    index = 0
    row = 0    
    if len(dfs) > 0:
        for x in dfs:
            column = 0
            data[row][0] = listOfCategory[index]
            row += 1
            if x.empty:
                processCol(x.columns, data[row])
                row += 1
                index += 1
                continue
            else:
                processCol(x.columns, data[row])
                row += 1
                for y in x.values:
                    for z in y:
                        # 먼저 z가 문자열인지 확인
                        if isinstance(z, str) and z.startswith("Unnamed:"):
                            data[row][column] = None
                        # z가 숫자형이고, np.nan인지 확인
                        elif isinstance(z, (int, float)) and np.isnan(z):
                            data[row][column] = None
                        else:
                            data[row][column] = z
                        column += 1
                    row += 1
                    column = 0
            if testSkipData(x):
                continue
            else:
                index += 1

def makeList(length, columnsCount):
    result = []
    for i in range(length):
        result.append([None for _ in range(columnsCount)])
    return result



# %%
if __name__ == "__main__":
    filePath = "C:\\Users\\young\\Downloads\\에이피엠_금융기관조회서\\AC3_AC3 KEB하나은행_APM(2023).pdf"
    saveFile = "C:\\Users\\young\\Downloads\\은행조회서_정리.xlsx"
    dfs = getDataFromPDF(filePath)
    # printPDFData(dfs)
    data = makeList(200, 15)
    processPDFData(dfs, data)
    #testPDFData(dfs)
    saveExcel(saveFile, data)
    print("Done")



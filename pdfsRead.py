#%%
import os
import tabula
import numpy as np
import psutil
import win32com.client as win

# %%
class PdfsReader:
    def __init__(self, sourcePsth) -> None:
        self.fileList = []
        self.data = self.makeList(500, 15)
        self.listOfCategory = [
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
        self.path = sourcePath
        self.dfs = None
        self.inputCount = 0
        self.xl = None
        self.wb = None
        self.ws = None
        if self.path is not None:
            self.getFileList(self.path)
            if len(self.fileList) == 0:
                print("file이 없습니다.")
            else:                
                if len(self.fileList) >= 1:
                    for p in  self.fileList: 
                        print(p)                     
                        self.dfs = self.getDataFromPDF(self.path + "\\" + p)
                        self.loadExcel()
                        self.processPDFData(self.dfs, self.data)
                        self.ws.Range(self.xl.Cells(1, 1), self.xl.Cells(len(self.data), len(self.data[0]))).Value = tuple(self.data) # type: ignore
                        sheetNameList = p.split("\\")
                        self.ws.Name = sheetNameList[len(sheetNameList)-1] # type: ignore
                        self.inputCount += 1

                    if self.wb is not None:
                        self.wb.SaveAs("C:\\Users\\young\\Downloads\\은행조회서_정리.xlsx")   
                        self.wb.Close()
            self.killExcel()

    def loadExcel(self):
        if self.xl is None:
            self.xl = win.gencache.EnsureDispatch("Excel.Application")
            self.xl.Visible = False
        if self.wb is None:
            self.wb = self.xl.Workbooks.Add()
        if self.inputCount < self.wb.Sheets.Count:
            self.ws = self.wb.Worksheets(self.inputCount + 1)
        else:
            self.ws = self.wb.Worksheets.Add()
            self.data = self.makeList(500, 15)

    def killExcel(self):
        if self.xl is not None:      
            self.xl.Quit()
        for proc in psutil.process_iter():
        # check whether the process name matches
            if proc.name() == "EXCEL.EXE":
                proc.kill()
    
    def getFileList(self, dir):
        answer = os.path.exists(sourcePath)

        if answer == True:
            self.fileList =  os.listdir(sourcePath)
        else:
            print(f"해당 폴더는 없습니다.")

    def getDataFromPDF(self, filePath):
        dfs = None
        if os.path.exists(filePath):
            print("File exists, proceeding with PDF processing.")
            # PDF 파일 읽기
            dfs = tabula.read_pdf(filePath, stream=True, pages='all') # type: ignore
            print(f"Number of dataframes: {len(dfs)}")
            return dfs
        else:
            print(f"File not found: {filePath}")
        return dfs  

    def makeList(self, length, columnsCount):
        result = []
        for i in range(length):
            result.append([None for _ in range(columnsCount)])
        return result
    
    def processCol(self, source, data):
        column = 0
        for ele in source:
            if isinstance(ele, str) and ele.startswith("Unnamed:"):
                data[column] = None
            else:
                data[column] = ele
            column += 1

    def testSkipData(self, dfs):
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

    def processPDFData(self, dfs, data):
        index = 0
        row = 1    
        if len(dfs) > 0:
            for x in dfs:
                column = 0
                if index < len(self.listOfCategory):
                    data[row][0] = self.listOfCategory[index]                    
                else:
                    data[row][0] = "Hello"
                row += 1
                if x.empty:
                    self.processCol(x.columns, data[row])
                    row += 2
                    index += 1
                    continue
                else:
                    self.processCol(x.columns, data[row])
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
                if self.testSkipData(x):
                    continue
                else:
                    index += 1
                row += 1
    
# %%

if __name__ == "__main__":
    
    
    sourcePath = "C:\\Users\\young\\Downloads\\에이피엠_금융기관조회서"

    answer = os.path.exists(sourcePath)

    if answer == True:
        print(f"해당 폴더는 존재합니다.")
    else:
        print(f"해당 폴더는 없습니다.")

    PdfsReader(sourcePath)
    print("Done")

# %%

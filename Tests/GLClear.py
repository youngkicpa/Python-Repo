import openpyxl

def gl_clear(file_path, target_path=r"C:\Users\young\Downloads\계정별원장_정리_SST.xlsx"):
    # 엑셀 파일 열기
    wb = openpyxl.load_workbook(file_path)
    
    sheet_Names = wb.sheetnames

    # 시트 지정
    tSheet = wb.create_sheet()  # 데이터를 저장할 시트
    
    # 변수 초기화
    start_row = 6  # Sheet2에 데이터를 저장할 시작 행
    code = "000"
    name = "계정과목"
    
    for i, x in enumerate(["계정코드", "계정과목", "날짜", "적요", "거래처번호", "거래처명", "차변", "대변", "잔액"], start=1):
        tSheet.cell(row=5, column=i, value=x)

    for sheet_Name in sheet_Names:
        # Sheet의 UsedRange처럼 모든 행을 순회
        for row in wb[sheet_Name].iter_rows(min_row=1, max_row=wb[sheet_Name].max_row, values_only=False):
            cell_1 = row[0].value  # A열 (회사명, 빈 값 등)
            cell_2 = row[1].value  # C열 (날짜, 전기이월 등)
            
            # 건너뛸 조건
            if (cell_1 is None and cell_2 != "전기이월") or \
            (isinstance(cell_2, str) and cell_2.replace(" ", "") in ["[월계]", "[누계]"]) or \
            (isinstance(cell_1, str) and  cell_1 == "날짜"):
                continue  # 다음 행으로 이동
            
            #if cell_1[1:4] == "회사명:주식회사 제이디":
            if isinstance(cell_1, str) and cell_1[0:4] == "회사명:": # type: ignore
                code = row[6].value[8:11]  # type: ignore # 9번째 문자부터 5글자 추출 (Python은 0부터 시작, 즉 index 8부터 13 전까지)
                name = row[6].value[13:35] # type: ignore
                continue
                
            # 데이터를 Sheet2로 복사
            for col_idx, cell in enumerate(row, start=1):  # A~마지막 열까지 반복
                tSheet.cell(row=start_row, column=1, value=code) # type: ignore
                tSheet.cell(row=start_row, column=2, value=name) # type: ignore
                tSheet.cell(row=start_row, column=col_idx + 2, value=cell.value)
            
            # "전기이월"이 포함된 경우 날짜를 "2024-01-01"로 변경
            if cell_2 == "전기이월":    
                tSheet.cell(row=start_row, column=1, value=code) # type: ignore
                tSheet.cell(row=start_row, column=2, value=name) # type: ignore
                tSheet.cell(row=start_row, column=3, value="01-01")  # C열 (3번째 열)에 값 설정

            start_row += 1  # 다음 행으로 이동
    
    # 파일 저장
    wb.save(target_path)
    print("엑셀 데이터 변환 완료!")

if __name__ == "__main__":
# 실행 예제
    file_path = r"C:\Users\young\Downloads\계정별원장 (SST).xlsx"
    gl_clear(file_path)

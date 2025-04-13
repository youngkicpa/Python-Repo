import openpyxl

def gl_clear_제이디(file_path):
    # 엑셀 파일 열기
    wb = openpyxl.load_workbook(file_path)
    
    # 시트 지정
    sheet1 = wb["Sheet1"]  # 원본 데이터가 있는 시트
    sheet2 = wb.create_sheet()  # 데이터를 저장할 시트
    
    # 변수 초기화
    start_row = 6  # Sheet2에 데이터를 저장할 시작 행
    code = "000"
    name = "계정과목"
    
    for i, x in enumerate(["계정코드", "계정과목", "날짜", "적요", "거래처번호", "거래처명", "차변", "대변", "잔액"], start=1):
        sheet2.cell(row=5, column=i, value=x)

    # Sheet1의 UsedRange처럼 모든 행을 순회
    for row in sheet1.iter_rows(min_row=1, max_row=sheet1.max_row, values_only=False):
        cell_1 = row[0].value  # A열 (회사명, 빈 값 등)
        cell_3 = row[2].value  # C열 (날짜, 전기이월 등)
        cell_4 = row[3].value  # D열 ([ 월 계 ], [ 누 계 ] 등)
        
        # 건너뛸 조건
        if (cell_1 is None and cell_4 != "전기이월") or \
           (cell_1 == "" and (isinstance(cell_4, str) and cell_4.replace(" ", "") in ["[월계]", "[누계]"])) or \
           (cell_3 == "날짜"):
            continue  # 다음 행으로 이동
        
        #if cell_1[1:4] == "회사명:주식회사 제이디":
        if isinstance(cell_1, str) and cell_1[0:4] == "회사명:": # type: ignore
            code = row[7].value[8:13]  # type: ignore # 9번째 문자부터 5글자 추출 (Python은 0부터 시작, 즉 index 8부터 13 전까지)
            name = row[7].value[15:35] # type: ignore
            continue
            
        # 데이터를 Sheet2로 복사
        for col_idx, cell in enumerate(row, start=1):  # A~마지막 열까지 반복
            sheet2.cell(row=start_row, column=col_idx, value=cell.value)
        
        # "전기이월"이 포함된 경우 날짜를 "2024-01-01"로 변경
        if cell_4 == "전기이월":    
            sheet2.cell(row=start_row, column=1, value=code) # type: ignore
            sheet2.cell(row=start_row, column=2, value=name) # type: ignore
            sheet2.cell(row=start_row, column=3, value="01-01")  # C열 (3번째 열)에 값 설정

        start_row += 1  # 다음 행으로 이동
    
    # 파일 저장
    wb.save(file_path)
    print("엑셀 데이터 변환 완료!")

if __name__ == "__main__":
# 실행 예제
    file_path = r"C:\Users\young\Downloads\계정별 원장_2024.12.31 기준_yskim.xlsx"
    gl_clear_제이디(file_path)

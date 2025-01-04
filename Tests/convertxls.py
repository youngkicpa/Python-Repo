import os
import win32com.client as win32

def convert_xls_to_xlsx(folder_path):
    # Excel 애플리케이션을 시작합니다
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False  # Excel 창을 표시하지 않음

    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.xls'):
                xls_path = os.path.join(root, file)
                xlsx_path = os.path.splitext(xls_path)[0] + '.xlsx'
                
                # .xls 파일 열기
                wb = excel.Workbooks.Open(xls_path)
                
                # .xlsx 형식으로 저장
                wb.SaveAs(xlsx_path, FileFormat=51)  # 51은 .xlsx 형식을 나타냅니다
                wb.Close()  # 작업 완료 후 워크북 닫기
                
                # 원본 .xls 파일 삭제
                os.remove(xls_path)
                print(f"Converted and deleted: {xls_path}")

    # Excel 애플리케이션 종료
    excel.Quit()

# 사용 예시
folder_path = r"C:\Users\young\OneDrive - 다산\019. DAA 폴더\2024 DAA_version_SK GAAP"
convert_xls_to_xlsx(folder_path)
print("Done")

import os



filename = "C:\\DataTest\\삼진엘앤디_분개장_상반기_yskim.xlsx"

listOfFile = os.listdir(os.path.dirname(filename))

if os.path.basename(filename) in listOfFile:
    print(f"{os.path.basename(filename)}이 존재합니다.")
else:
    print(f"{os.path.basename(filename)}이 없습니다.")



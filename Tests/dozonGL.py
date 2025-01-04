import pandas as pd

class DozonGLDataframe:
    def __init__(self, table):
        self.title = table[0]
        self.data = pd.DataFrame(table[1:], columns=self.title)
    
    def getTriablBalance(self):
        # 합계잔액시산표를 만드는 코드
        trialBalance = self.data[['계정코드', '계정명', '차변', '대변']].groupby(['계정코드', '계정명']).sum()
        trialBalance['잔액'] = trialBalance['차변'] - trialBalance['대변']
        # 각 계정코드와 계정명의 빈도를 계산
        frequency = self.data.groupby(['계정코드', '계정명']).size().reset_index(name='빈도')

        # trialBalance에 빈도 열 추가
        return trialBalance.merge(frequency, on=['계정코드', '계정명'], how='left')

    def getTrialBalanceKSsystem(self):
        # 합계잔액시산표를 만드는 코드
        trialBalance = self.data[['계정과목', '차변', '대변']].groupby(['계정과목']).sum()
        trialBalance['잔액'] = trialBalance['차변'] - trialBalance['대변']
        # 각 계정코드와 계정명의 빈도를 계산
        frequency = self.data.groupby(['계정과목']).size().reset_index(name='빈도')

        # trialBalance에 빈도 열 추가
        return trialBalance.merge(frequency, on=['계정과목'], how='left')
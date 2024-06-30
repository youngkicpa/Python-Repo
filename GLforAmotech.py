from MyPPrint import *

class Amotech:
    def __init__(self, data) -> None:
        self.data = data
        self.trialBalance = {}

    def testforTitles(self, row):
        if row[0] == "회계일":
            return True
        else:
            return False
        
    def getTransaction(self, data):
        start = False
        trannsaction = {}
        for x in self.data:
            if self.testforTitles(x):
                start = True
                continue
            if not start:
                continue
            else:
                continue        
        return None
        
    def getTrialBalance(self):
        start = False
        for x in self.data:
            if self.testforTitles(x):
                start = True
                continue
            if not start:
                continue
            else:
                if (x[3], x[4]) not in self.trialBalance:
                    self.trialBalance[(x[3], x[4])] = [0, 0, 0]
                    self.trialBalance[(x[3], x[4])][0] = 1
                    self.trialBalance[(x[3], x[4])][1] = x[6]
                    self.trialBalance[(x[3], x[4])][2] = x[7]
                else:
                    self.trialBalance[(x[3], x[4])][0] += 1
                    self.trialBalance[(x[3], x[4])][1] += x[6]
                    self.trialBalance[(x[3], x[4])][2] += x[7]
    
    def printTrialBalance(self):
        sumCount = 0
        sumDebit = 0
        sumCredit = 0
        header = ["계정코드", "계정명", "빈도수", "차변", "대변"]
        lengths = [10, 50, -10, -20, -20]
        strings = (str(header[0]), header[1], header[2], header[3])
        myPPrint(lengths, strings)
        for k, v in sorted(self.trialBalance.items()):
            code = k[0]
            name = k[1]
            strings = (code, name, "{:>10,.0f}".format(v[0]), "{:>20,.0f}".format(v[1]), "{:>20,.0f}".format(v[2]) )
            sumCount += v[0]
            sumDebit += v[1]
            sumCredit += v[2]
            myPPrint(lengths, strings)

        lengths.append(-20)
        strings = ["합계: ", "", "{:>10,.0f}".format(sumCount), "{:>20,.0f}".format(sumDebit), "{:>20,.0f}".format(sumCredit), "{:>20,.0f}".format(sumDebit-sumCredit)]
        myPPrint(lengths, strings)

        

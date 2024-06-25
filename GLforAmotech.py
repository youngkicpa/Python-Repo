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
                    self.trialBalance[(x[3], x[4])] = [0, 0]
                    self.trialBalance[(x[3], x[4])][0] = x[6]
                    self.trialBalance[(x[3], x[4])][1] = x[7]
                else:
                    self.trialBalance[(x[3], x[4])][0] += x[6]
                    self.trialBalance[(x[3], x[4])][1] += x[7]
    
    def printTrialBalance(self):
        sumDebit = 0
        sumCredit = 0
        header = ["계정코드:계정명", "차변", "대변"]
        lengths = [50, -20, -20]
        strings = (str(header[0]), header[1], header[2] )
        myPPrint(lengths, strings)
        for k, v in sorted(self.trialBalance.items()):
            strings = (str(k), "{:>20,.0f}".format(v[0]), "{:>20,.0f}".format(v[0]) )
            sumDebit += v[0]
            sumCredit += v[1]
            myPPrint(lengths, strings)

        lengths.append(-20)
        strings = ["합계: ", "{:>20,.0f}".format(sumDebit), "{:>20,.0f}".format(sumCredit), "{:>20,.0f}".format(sumDebit-sumCredit)]
        myPPrint(lengths, strings)

        

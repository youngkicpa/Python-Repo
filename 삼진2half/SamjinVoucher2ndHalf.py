class SamjinVoucher:
    def __init__(self) -> None:
        self.no = ""
        self.debit =  { "amounts": [], "accounts":[], "codes":[] }
        self.credit = { "amounts": [], "accounts":[], "codes":[] }
        self.debitSum = 0
        self.creditSum = 0
        self.amountTest = False
        self.codeColNo = 1
        self.accountColNo = 2
        self.debitColNo = 6
        self.creditColNo = 8
        self.voucherNoCol = 3

    def Add(self, row):
        if self.no == "":
            self.no = row[self.voucherNoCol]
        if row[self.debitColNo] == 0:
            self.credit["amounts"].append(row[self.creditColNo])
            self.credit["codes"].append(row[self.codeColNo])
            self.credit["accounts"].append(row[self.accountColNo])
            self.creditSum += row[self.creditColNo]
        else:
            self.debit["amounts"].append(row[self.debitColNo])
            self.debit["codes"].append(row[self.codeColNo])
            self.debit["accounts"].append(row[self.accountColNo])
            self.debitSum += row[self.debitColNo]

    def TestAmounts(self):
        return self.debitSum == self.creditSum
    
    def ToList(self):
        result = []
        for i in range(len(self.debit["amounts"])):
            result.append((self.no, self.debit["codes"][i], self.debit["accounts"][i], self.debit["amounts"][i], 0))
        for i in range(len(self.credit["amounts"])):
            result.append((self.no, self.credit["codes"][i], self.credit["accounts"][i], 0, self.credit["amounts"][i]))
        return result
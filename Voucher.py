class Voucher:
    def __init__(self) -> None:
        self.no = ""
        self.debit =  { "amounts": [], "accounts":[], "codes":[], "prepared":[], "dep":[] }
        self.credit = { "amounts": [], "accounts":[], "codes":[], "prepared":[], "dep":[] }
        self.debitSum = 0
        self.creditSum = 0
        self.amountTest = False

    def Add(self, row):
        if self.no == "":
            self.no = row[1][0:13]
        if row[6] == 0:
            self.credit["amounts"].append(row[7])
            self.credit["codes"].append(row[3])
            self.credit["accounts"].append(row[4])
            self.credit["prepared"].append(row[14])
            self.credit["dep"].append(row[15])
            self.creditSum += row[7]
        else:
            self.debit["amounts"].append(row[6])
            self.debit["codes"].append(row[3])
            self.debit["accounts"].append(row[4])
            self.debit["prepared"].append(row[14])
            self.debit["dep"].append(row[15])
            self.debitSum += row[6]

    def TestAmounts(self):
        return self.debitSum == self.creditSum
    
    def ToList(self):
        result = []
        for i in range(len(self.debit["amounts"])):
            result.append((self.no, self.debit["codes"][i], self.debit["accounts"][i], self.debit["amounts"][i], 0))
        for i in range(len(self.credit["amounts"])):
            result.append((self.no, self.credit["codes"][i], self.credit["accounts"][i], 0, self.credit["amounts"][i]))
        return result
    
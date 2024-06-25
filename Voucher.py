class Voucher:
    def __init__(self) -> None:
        self.no = ""
        self.debit = { "amounts": [], "accounts":[], "codes":[] }
        self.credit = { "amounts": [], "accounts":[], "codes":[] }
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
            self.creditSum += row[7]
        else:
            self.debit["amounts"].append(row[6])
            self.debit["codes"].append(row[3])
            self.debit["accounts"].append(row[4])
            self.debitSum += row[6]

    def TestAmounts(self):
        return self.debitSum == self.creditSum
    
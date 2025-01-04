from SamjinVoucher import *

class SamjinVouchers:
    def __init__(self):
        self.vouchers = []
        self.trialBalance = {}

    def testforTitles(self, row):
        if row[0] == "회계일":
            return True
        else:
            return False   

    def getVouchers(self, data):
        previousNo = ""
        start = False
        currentVoucher = SamjinVoucher()
        for i, row in enumerate(data):
            if self.testforTitles(row):
                start = True
                continue
            if not start:
                continue
            else:
                if row[12] == previousNo:
                        currentVoucher.Add(row)
                else:
                        self.vouchers.append(currentVoucher)
                        currentVoucher = SamjinVoucher()
                        currentVoucher.Add(row)
                        previousNo = row[12]
                if i == len(data) - 1:
                    self.vouchers.append(currentVoucher)
                     
    def testVoucherAmount(self):
        for voucher in self.vouchers:
            if not voucher.TestAmounts():
                print(f"{voucher.no}\t {voucher.debitSum}\t{voucher.creditSum}\t{voucher.TestAmounts()}")
        print("\n\n전표의 차대변 합계 검증이 끝났습니다.\n\n")

    def testVoucherSales(self, voucher):
        if len(voucher.credit["codes"]) == 0:
            return False
        for x in voucher.credit["codes"]:
            if 51000101 <= int(x) <= 51000502:
                return True 
        return False

    def testVoucherMinusSales(self, voucher):
        if len(voucher.credit["codes"]) == 0:
            return False
        for index, x in enumerate(voucher.credit["codes"]):
            if 51000101 <= int(x) <= 51000502 and voucher.credit["amounts"][index] < 0:
                return True 
        return False

    def getFiltered(self, condition):
        filtered = []
        for voucher in self.vouchers:
            if condition(voucher):
                filtered.append(voucher.no)

        return filtered

    def salesTransactions(self):
        debitInfo = {}
        creditInfo = {}
        for voucher in self.vouchers:
            if self.testVoucherSales(voucher):
                for index, d in enumerate(voucher.debit["accounts"]):
                    if d in debitInfo:
                        debitInfo[d] += voucher.debit["amounts"][index]
                    else:
                        debitInfo[d] = voucher.debit["amounts"][index]

                for index, c in enumerate(voucher.credit["accounts"]):
                    if c in creditInfo:
                        creditInfo[c] += voucher.credit["amounts"][index]
                    else:
                        creditInfo[c] = voucher.credit["amounts"][index]

        print("차변")
        for key, value in debitInfo.items():
            print(f"{key}\t{value}")
        print("대변")
        for key, value in creditInfo.items():
            print(f"{key}\t{value}")

    def minusSalesTransactions(self):
        debitInfo = {}
        creditInfo = {}
        count = 0
        filtered = []
        for voucher in self.vouchers:
            if self.testVoucherMinusSales(voucher):
                count += 1
                filtered.append(voucher)
                for index, d in enumerate(voucher.debit["accounts"]):
                    if d in debitInfo:
                        debitInfo[d] += voucher.debit["amounts"][index]
                    else:
                        debitInfo[d] = voucher.debit["amounts"][index]

                for index, c in enumerate(voucher.credit["accounts"]):
                    if c in creditInfo:
                        creditInfo[c] += voucher.credit["amounts"][index]
                    else:
                        creditInfo[c] = voucher.credit["amounts"][index]

        print(f"{count}개의 전표가 있습니다.")
        print("차변")
        for key, value in debitInfo.items():
            print(f"{key}\t{value}")
        print("대변")
        for key, value in creditInfo.items():
            print(f"{key}\t{value}")
        return filtered

    def getVouchersAmounts(self):
        count = 0
        count_minus = 0
        resultList = []
        result = { 
            "백억초과": 0,
            "백억이하": 0,
            "칠십오억이하": 0,
            "오십억이하": 0,
            "이십오억이하": 0,
            "십억이하": 0,
            "일억이하":0,
            "(-)전표": 0
        }
        for voucher in self.vouchers:
            count += 1
            match voucher.creditSum:
                case n if n < 0:
                    result["(-)전표"] += 1
                    if n < -100000000:
                        count_minus += 1
                        #resultList.extend(voucher.ToList())
                case n if 0 <= n <= 100000000:
                    result["일억이하"] += 1
                case n if 100000000 < n <= 1000000000:
                    result["십억이하"] += 1
                case n if 1000000000 < n <= 2500000000:
                    result["이십오억이하"] += 1
                case n if 2500000000 < n <= 5000000000:
                    result["오십억이하"] += 1
                case n if 5000000000 < n <= 7500000000:
                    result["칠십오억이하"] += 1
                case n if 7500000000 < n <= 10000000000:
                    result["백억이하"] += 1
                case n if 10000000000 < n:
                    result["백억초과"] += 1
                    resultList.extend(voucher.ToList())

        print(f"\n전표의 총갯수는 : {count}")
        print(f"(-)1억미만전표의 갯수는: {count_minus}")
        for key, value in result.items():
            print(f"{key}: \t {value:>7}\t개\t {value/count*100 if count != 0 else 0.00:>6.2f}%")
        print("\n")
        return resultList    

    def getTrialBalance(self):
        debitsum = 0
        creidtsum = 0
        for x in self.vouchers:
            for i in range(len(x.debit["codes"])):
                if (x.debit["codes"][i], x.debit["accounts"][i]) not in self.trialBalance.keys():
                    self.trialBalance[(x.debit["codes"][i], x.debit["accounts"][i])] = [0, 0, 0]
                    self.trialBalance[(x.debit["codes"][i], x.debit["accounts"][i])][0] = x.debit["amounts"][i]
                    debitsum += x.debit["amounts"][i]
                    self.trialBalance[(x.debit["codes"][i], x.debit["accounts"][i])][2] = 1
                else:
                    self.trialBalance[(x.debit["codes"][i], x.debit["accounts"][i])][0] += x.debit["amounts"][i]
                    debitsum += x.debit["amounts"][i]
                    self.trialBalance[(x.debit["codes"][i], x.debit["accounts"][i])][2] += 1

            for i in range(len(x.credit["codes"])):
                if (x.credit["codes"][i], x.credit["accounts"][i]) not in self.trialBalance.keys():
                    self.trialBalance[(x.credit["codes"][i], x.credit["accounts"][i])] = [0, 0, 0]
                    self.trialBalance[(x.credit["codes"][i], x.credit["accounts"][i])][1] = x.credit["amounts"][i]
                    creidtsum += x.credit["amounts"][i]
                    self.trialBalance[(x.credit["codes"][i], x.credit["accounts"][i])][2] = 1 
                else:
                    self.trialBalance[(x.credit["codes"][i], x.credit["accounts"][i])][1] += x.credit["amounts"][i]
                    creidtsum += x.credit["amounts"][i]
                    self.trialBalance[(x.credit["codes"][i], x.credit["accounts"][i])][2] += 1 
        
        print(f"차변합계 : {debitsum}, 대변합계: {creidtsum}")

    def getTrialBalanceTuple(self):
        result = []
        result.append(("계정코드", "계정명", "차변", "대변", "횟수"))
        for key, value in sorted(self.trialBalance.items()):
            result.append((key[0], key[1], value[0], value[1], value[2]))
        
        return result
from Voucher import *

class Vouchers:
    def __init__(self):
        self.vouchers = [] 

    def testforTitles(self, row):
        if row[0] == "회계일":
            return True
        else:
            return False   

    def getVouchers(self, data):
        previousNo = ""
        start = False
        currentVoucher = Voucher()
        for row in data:
            if self.testforTitles(row):
                start = True
                continue
            if not start:
                continue
            else:
                if row[1][0:13] == previousNo:
                        currentVoucher.Add(row)
                else:
                        self.vouchers.append(currentVoucher)
                        currentVoucher = Voucher()
                        currentVoucher.Add(row)
                        previousNo = row[1][0:13]
                     
    def testVoucherAmount(self):
        for voucher in self.vouchers:
            if not voucher.TestAmounts():
                print(f"{voucher.no}\t {voucher.debitSum}\t{voucher.creditSum}\t{voucher.TestAmounts()}")
        print("\n\n전표의 차대변 합계 검증이 끝났습니다.\n\n")

    def testVoucherSales(self, voucher):
        if len(voucher.credit["codes"]) == 0:
            return False
        for x in voucher.credit["codes"]:
            if 4110000 <= int(x) <= 4200000:
                return True 
        return False

    def testVoucherMinusSales(self, voucher):
        if len(voucher.credit["codes"]) == 0:
            return False
        for index, x in enumerate(voucher.credit["codes"]):
            if 4110000 <= int(x) <= 4200000 and voucher.credit["amounts"][index] < 0:
                return True 
        return False

    def getFiltered(self, condition):
        filtered = []
        for voucher in self.vouchers:
            if condition(voucher):
                filtered.append(voucher)

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

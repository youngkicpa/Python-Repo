'''
 한글을 사용해도 프린트가 예쁘게 되도록 print를 수정한 것이다.
'''

def getUTFNo(string):
    countUTF = [ True for x in string if ord(x) > 127].count(True)
    return countUTF

def getStrForPrint(length, string):
    '''
    이 함수는 문자열의 길이가 표시되어야 하는 길이보다 길 때 조정해주기 위한 것이다.
    '''
    result = ''
    length *= -1 if length < 0 else 1
    count = 0
    for x in string if type(string) is str else str(string):
        count += 2 if ord(x) > 127 else 1
        if count <= length:
            result += x
        else:
            return result
    return result
    
def myPPrint(lengths, strings, delimeter='  '):
    for x in zip(lengths, strings):
        result = getStrForPrint(x[0], x[1])
        count = getUTFNo(result)
        resultLen = abs(x[0]) - count
        if x[0] < 0:
            printformat = f"{{:>{resultLen}}}"
        else:          
             printformat = f"{{:<{resultLen}}}"
        print(printformat.format(result), end=delimeter)
    print()

def printdict(head, pair, delimeter):
    result = getStrForPrint(pair[0], pair[1])
    count = getUTFNo(result)
    resultLen = abs(pair[0])-count
    if head:
        printformat = "{0:<{1}}"
    else:
        printformat = "{0:>{1}}" if pair[0] < 0 else "{0:{1}}"
    print(printformat.format(result,  resultLen), end=delimeter)

def dictPPrint(head, lengths, data, delimeter='  '):
    '''
      헤더는 첫번째 Dictionary에만 포함을 하면 된다. 이때 keys()에 대해서만 적용한다.
    '''
    if head:
        for x in zip(lengths, data.keys()):
            printdict(head, x, delimeter)
        
    # value에는 헤더가 True일 필요가 없다. 어떤 경우에도
    for x in zip(lengths, data.values()):
        printdict(False, x, delimeter)
    print()
from xlrd import open_workbook
from xlutils.copy import copy
from re import *
from datetime import *
from dateutil.relativedelta import *

# 날짜 일괄 변경
def DateEdit():
    inputDate = input("날짜(YYYY.MM.DD)를 입력하세요. : ")

    for i in range(1, num_rows):
        writeSheet.write(i, 0, inputDate)

    print("날짜를 수정했습니다.")

# 필요한 부분 패턴화하여 매칭
def PatternMatch(data):
    # [연장01-02  #   /04월분    #   ]2020.04.30만료   #
    patOne = search('\[연장\d\d-\d\d', data)
    patTwo = search('/\d\d월분', data)
    patThree = search('\]\d\d\d\d.\d\d.\d\d만료', data)

    matched = [patOne.group(), patTwo.group(), patThree.group()]

    bfPat1 = patOne.start() + 3
    afPat1 = patOne.end()
    bfPat2 = patTwo.start() + 1
    afPat2 = patTwo.end() - 2
    bfPat3 = patThree.start() +1
    afPat3 = patThree.end() - 2

    index = [bfPat1, afPat1, bfPat2, afPat2, bfPat3, afPat3]

    return [matched, index]

# 문자열 자르기 함수
def StringIndexing(data):
    matched = PatternMatch(data)[0]

    date1 = matched[0][3:8]     # YY-MM
    date2 = matched[1][1:3]     # MM
    date3 = matched[2][1:11]    # YYYY.MM.DD

    strDates = [date1, date2, date3]

    return strDates


# 날짜-문자열변환 및 1달 합
def DateTimeConvertPlusConvert(data):
    strDates = StringIndexing(data)
    date1 = datetime.strptime(strDates[0], '%y-%m').date() + relativedelta(months=1)
    date2 = datetime.strptime(strDates[1], '%m').date() + relativedelta(months=1)
    date3 = datetime.strptime(strDates[2], '%Y.%m.%d').date() + relativedelta(months=1)

    dates = [date1.strftime("%y-%m"), date2.strftime("%m"), date3.strftime("%Y.%m.%d")]

    return dates

# 적요 수정
def BriefEdit():
    for i in range(1, num_rows, 2):
        data = readSheet.cell_value(i, 10)

        dates = DateTimeConvertPlusConvert(data)
        index = PatternMatch(data)[1]

        update = data[:index[0]] + dates[0] + data[index[1]:index[2]] + dates[1] + data[index[3]:index[4]] + dates[2] + data[index[5]:]
        writeSheet.write(i, 10, update)

    print("적요를 수정했습니다.")


# ---------------메인함수-------------------

# file = "과제파일.xls"
inputFile = input("파일명(확장자명 제외)을 입력하세요. : ")
inputFile += ".xls"

readWorkBook = open_workbook(inputFile, formatting_info=True)
readSheet = readWorkBook.sheet_by_index(0)

writeWorkBook = copy(readWorkBook)
writeSheet = writeWorkBook.get_sheet(0)

num_rows = readSheet.nrows

# 날짜 변경
DateEdit()

# 적요 변경
BriefEdit()

# 파일 저장
saveFile = input("저장할 파일명(확장자명 제외)을 입력하세요. : ")
saveFile += '.xls'
writeWorkBook.save(saveFile)
print("완료되었습니다. 프로그램을 종료합니다.")

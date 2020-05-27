from xlrd import open_workbook
from xlutils.copy import copy
from re import *
from datetime import *
from dateutil.relativedelta import *
from os import *
from io import *


# 파일 존재 여부 확인
def IsFileValid():
    while True:
        filePath = input("엑셀파일(.xls)의 경로를 입력하세요. : ")
        if path.isfile(filePath):
            return filePath
        else:
            print("존재하지않는 파일입니다. 다시 확인해주세요.")


# 입력형식이 맞는 유효한 날짜형식인지 확인
def IsInputDateValid(date):
    try:
        datetime.strptime(date, '%Y.%m.%d')
        return True
    except ValueError:
        print("날짜형식 또는 입력형식이 맞지 않습니다. 다시 입력해주세요.")
        return False


# 날짜 일괄 변경
def DateEdit():
    while True:
        inputDate = input("날짜(YYYY.MM.DD)를 입력하세요. : ")
        if IsInputDateValid(inputDate):
            for i in range(1, num_rows):
                data = readSheet.cell_value(i, 0)
                if search('\d\d\d\d.\d\d.\d\d', data):
                    writeSheet.write(i, 0, inputDate)
            print("날짜를 수정했습니다.")
            break


# 적요 데이터 패턴
def DataPatterns(data):
    return [search('\[연장\d\d-\d\d', data), search('/\d\d월분', data), search('\]\d\d\d\d.\d\d.\d\d만료', data)]


# 데이터 오류 확인
def DataErrorCheck(data):
    if (DataPatterns(data)[0] is None) or (DataPatterns(data)[1] is None) or (DataPatterns(data)[2] is None):
        return False
    elif StringToDatetime(data) is False:
        return False
    else:
        return True


# 공백 없애기
def DataTrim(data):
    dataEnd = search('만료', data)
    index = dataEnd.end()
    result = data[:index].replace(" ", "") + data[index:]
    return result


# 필요한 부분 패턴화하여 매칭
def PatternMatch(data):
    # [연장01-03*(2인)/05월분]2020.05.30만료
    # [연장01-02  #   /04월분    #   ]2020.04.30만료   #
    patOne = DataPatterns(data)[0]
    patTwo = DataPatterns(data)[1]
    patThree = DataPatterns(data)[2]

    matched = [patOne.group(), patTwo.group(), patThree.group()]

    bfPat1 = patOne.start() + 3
    afPat1 = patOne.end()
    bfPat2 = patTwo.start() + 1
    afPat2 = patTwo.end() - 2
    bfPat3 = patThree.start() + 1
    afPat3 = patThree.end() - 2

    index = [bfPat1, afPat1, bfPat2, afPat2, bfPat3, afPat3]

    return [matched, index]


# 문자열 자르기 함수
def StringIndexing(data):
    matched = PatternMatch(data)[0]

    date1 = matched[0][3:8]  # YY-MM
    date2 = matched[1][1:3]  # MM
    date3 = matched[2][1:11]  # YYYY.MM.DD

    strDates = [date1, date2, date3]

    return strDates


# 문자열 -> 날짜 변환
def StringToDatetime(data):
    strDates = StringIndexing(data)

    try:
        date1 = datetime.strptime(strDates[0], '%y-%m')
        date2 = datetime.strptime(strDates[1], '%m')
        date3 = datetime.strptime(strDates[2], '%Y.%m.%d')

    except ValueError:
        return False

    return [date1, date2, date3]


# 날짜 1달 증가 및 문자열 변환
def Plus1MonthAndToString(data):
    date1 = StringToDatetime(data)[0].date() + relativedelta(months=1)
    date2 = StringToDatetime(data)[1].date() + relativedelta(months=1)
    date3 = StringToDatetime(data)[2].date() + relativedelta(months=1)

    date3Month = date3.strftime("%Y.%m.%d")[5:7]
    date3Day = date3.strftime("%Y.%m.%d")[8:]
    if date3Month == '03' and date3Day == '29':
        date3 += timedelta(days=1)
    if date3Month == '03' and date3Day == '28':
        date3 += timedelta(days=2)

    dates = [date1.strftime("%y-%m"), date2.strftime("%m"), date3.strftime("%Y.%m.%d")]

    return dates


# 입력 데이터에 오류가 있을 시, 로그 파일 생성
def LogFile(row):
    if not path.exists('logs'):
        makedirs('logs')
    if not path.isfile(".\\logs\\log.txt"):
        logFile = open(".\\logs\\log.txt", "w", encoding='utf-8')
        logFile.write("파일의 데이터를 수정하는 중, 데이터에 오류가 발견되었습니다.\n")
        logFile.write("다음 부분의 데이터를 다시 확인하고 이 파일은 삭제하세요.\n")
        logFile.write("-------------------------------------------------------------------\n")
        logFile.close()
    logFile = open(".\\logs\\log.txt", "a", encoding='utf-8')
    logFile.write(str(row + 1) + '행의 적요 데이터에 문제가 있습니다.\n')
    logFile.close()


# 적요 수정
def BriefEdit():
    for i in range(1, num_rows):
        data = readSheet.cell_value(i, 10)
        if search('\[연장', data) is not None:
            index = search('\[연장', data).start()
            trimedData = DataTrim(data[index:])
            if DataErrorCheck(trimedData):
                front = data[:index]
                dates = Plus1MonthAndToString(trimedData)
                index = PatternMatch(trimedData)[1]
                tail = index[4] + len(dates[2])
                update = front + trimedData[:index[0]] + dates[0] + trimedData[index[1]:index[2]] + dates[1] \
                         + trimedData[index[3]:index[4]] + dates[2] + trimedData[tail:]
                writeSheet.write(i, 10, update)
            else:
                LogFile(i)

    print("적요를 수정했습니다.")


# 파일 저장 후 엑셀파일 열기
def SaveFileAndOpen():
    saveFile = input("저장할 파일명(확장자명 제외)을 입력하세요. : ")
    saveFile += '.xls'
    if not path.exists('saved'):
        makedirs('saved')
    filePath = ".\\saved\\" + saveFile
    logPath = ".\\logs\\log.txt"
    writeWorkBook.save(filePath)
    print("saved 폴더에 저장되었습니다. 프로그램을 종료하고 엑셀 파일과 로그 파일을 엽니다.")
    system('start excel.exe "%s"' % (filePath))
    system('start notepad.exe "%s"' % (logPath))


# ---------------메인함수-------------------

# 파일명 확인
inputPath = IsFileValid()

readWorkBook = open_workbook(inputPath, formatting_info=True)
readSheet = readWorkBook.sheet_by_index(0)

writeWorkBook = copy(readWorkBook)
writeSheet = writeWorkBook.get_sheet(0)

num_rows = readSheet.nrows

# 날짜 변경
DateEdit()

# 적요 변경
BriefEdit()

# 파일 저장 후 열기
SaveFileAndOpen()

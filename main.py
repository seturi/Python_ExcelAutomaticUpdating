from xlrd import open_workbook
from xlutils.copy import copy
from re import *
from datetime import *
from dateutil.relativedelta import *
from os import *

# 파일 존재 여부 확인
def IsFileValid():
    while True:
        filePath = input("엑셀파일(.xls)의 경로를 입력하세요. : ")
        if path.isfile(filePath):
            return filePath
        else:
            print("존재하지않는 파일입니다. 다시 확인해주세요.")
            continue

# 입력형식이 맞는 유효한 날짜형식인지 확인
def IsDateValid(inputDate):
    try:
        datetime.strptime(inputDate, '%Y.%m.%d')
        return True
    except ValueError:
        print("날짜형식 또는 입력형식이 맞지 않습니다. 다시 입력해주세요.")
        return False

# 날짜 일괄 변경
def DateEdit():
    while True:
        inputDate = input("날짜(YYYY.MM.DD)를 입력하세요. : ")
        if IsDateValid(inputDate):
            for i in range(1, num_rows):
                writeSheet.write(i, 0, inputDate)
            print("날짜를 수정했습니다.")
            break
        else:
            continue

# 공백 없애기
def DataTrim(data):
    dataEnd = search('만료', data)
    index = dataEnd.end()
    result = data[:index].replace(" ", "") + data[index:]
    return result

# 필요한 부분 패턴화하여 매칭
def PatternMatch(data):
    # (CMS)[연장01-03*(2인)/05월분]2020.05.30만료
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

    date3Month = date3.strftime("%Y.%m.%d")[5:7]
    date3Day = date3.strftime("%Y.%m.%d")[8:]
    if date3Month == '03' and date3Day == '29':
        date3 += timedelta(days=1)
    if date3Month == '03' and date3Day == '28':
        date3 += timedelta(days=2)

    dates = [date1.strftime("%y-%m"), date2.strftime("%m"), date3.strftime("%Y.%m.%d")]

    return dates

# 적요 수정
def BriefEdit():
    for i in range(1, num_rows):
        data = readSheet.cell_value(i, 10)
        if data and search('\(CMS\)', data):
            trimedData = DataTrim(data)
            dates = DateTimeConvertPlusConvert(trimedData)
            index = PatternMatch(trimedData)[1]
            tail = index[4] + len(dates[2])
            update = trimedData[:index[0]] + dates[0] + trimedData[index[1]:index[2]] + dates[1] + trimedData[index[3]:index[4]] + dates[2] + trimedData[tail:]
            writeSheet.write(i, 10, update)

    print("적요를 수정했습니다.")

# 파일 저장 후 엑셀파일 열기
def SaveFileAndOpen():
    saveFile = input("저장할 파일명(확장자명 제외)을 입력하세요. : ")
    saveFile += '.xls'
    if not path.exists('saved'):
        makedirs('saved')
    filePath = str(getcwd()) + "\\saved\\" + saveFile
    writeWorkBook.save(filePath)
    print("saved 폴더에 저장되었습니다. 프로그램을 종료하고 파일을 엽니다.")
    system('start excel.exe "%s"' % (filePath))


# ---------------메인함수-------------------

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

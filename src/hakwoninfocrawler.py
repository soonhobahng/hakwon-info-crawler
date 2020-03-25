from urllib.parse import urlencode
from openpyxl import Workbook
import datetime
import json
import requests

now = datetime.datetime.now()
areaUrls = [
    {'url': 'https://hakwon.sen.go.kr', 'areaNm': '서울'},
    {'url': 'https://hakwon.pen.go.kr', 'areaNm': '부산'},
    {'url': 'https://hakwon.dge.go.kr', 'areaNm': '대구'},
    {'url': 'https://hakwon.ice.go.kr', 'areaNm': '인천'},
    {'url': 'https://hakwon.gen.go.kr', 'areaNm': '광주'},
    {'url': 'https://hakwon.dje.go.kr', 'areaNm': '대전'},
    {'url': 'https://hakwon.use.go.kr', 'areaNm': '울산'},
    {'url': 'https://hakwon.sje.go.kr', 'areaNm': '세종시'},
    {'url': 'https://hakwon.goe.go.kr', 'areaNm': '경기도'},
    {'url': 'https://hakwon.kwe.go.kr', 'areaNm': '강원도'},
    {'url': 'https://hakwon.cbe.go.kr', 'areaNm': '충북'},
    {'url': 'https://hakwon.cne.go.kr', 'areaNm': '충남'},
    {'url': 'https://hakwon.jbe.go.kr', 'areaNm': '전북'},
    {'url': 'https://hakwon.jne.go.kr', 'areaNm': '전남'},
    {'url': 'https://hakwon.gbe.kr', 'areaNm': '경북'},
    {'url': 'https://hakwon.gne.go.kr', 'areaNm': '경남'},
    {'url': 'https://hakwon.jje.go.kr', 'areaNm': '제주'},
]

zoneCodes = []
searchParams = {}

# Print iterations progress
def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print('\r%s |%s| %s%% %s' % (prefix, bar, percent, suffix), end = printEnd)
    # Print New Line on Complete
    if iteration == total:
        print()

def hakwondata(zcode, neisUrl, areaNm, cookies):
    book = Workbook()
    urllist = []
    excelList = []
    teacherList = []
    totalPage = 1
    zoneNm = ""

    searchParams["searchZoneCode"] = zcode

    ## to get total count
    response = requests.post(neisUrl + '/scs_ica_cr91_005.ws', data=json.dumps(searchParams), cookies=cookies)
    jsonObj = json.loads(response.content)

    ## total count
    totalCnt = int(jsonObj["resultSVO"]["totalCount"])
    print(totalCnt, "records ")

    if totalCnt > 1000:
        totalPage, therest = divmod(totalCnt, 1000)
        if therest > 0:
            totalPage = totalPage + 1
        searchParams["pageSize"] = 1000
    elif totalCnt < 1:
        print("No records found")
        return
    else:
        totalPage = 1
        searchParams["pageSize"] = str(totalCnt)

    print("Total ", totalPage, " pages...")

    zoneNm = findZoneName(zcode)

    for pageIndex in range(0, totalPage, 1):
        print(pageIndex + 1, " page crawling... ")
        ## post again
        searchParams["pageIndex"] = str(pageIndex + 1)
        response = requests.post(neisUrl + '/scs_ica_cr91_005.ws', data=json.dumps(searchParams), cookies=cookies)
        jsonObj = json.loads(response.content)
        hakwonlist = jsonObj["resultSVO"]["hesIcaCr91M00DVO"]

        seq = (pageIndex) * int(searchParams["pageSize"]) + 1
        teacher_params = {}

        curAcaAsnum = ""

        progressIndex = 0
        maxProgressIndex = len(hakwonlist)
        printProgressBar(progressIndex, maxProgressIndex, prefix='Progress:', suffix='Complete', length=50)

        for onehakwon in hakwonlist:
            leSbjtNm = ""
            if isinstance(onehakwon["leSbjtNm"], list):
                leSbjtNm = onehakwon["leSbjtNm"][0]
            else:
                leSbjtNm = onehakwon["leSbjtNm"]

            lessonPeriod = onehakwon["leMms"] + '개월 ' + onehakwon["lePrdDds"] + '일'
            excelData = (strNow, seq, onehakwon["zoneNm"], onehakwon["acaNm"], onehakwon["gmNm"], leSbjtNm, onehakwon["faTelno"], onehakwon["totalJuso"], onehakwon["toforNmprFgr"], lessonPeriod,
                         onehakwon["totLeTmMmFgr"], onehakwon["thccSmtot"], onehakwon["thccAmt"], onehakwon["etcExpsSmtot"])

            if curAcaAsnum != onehakwon["acaAsnum"]:
                curAcaAsnum = onehakwon["acaAsnum"]
                ## 강사 정보 추출
                teacher_params["juOfcdcCode"] = onehakwon["juOfcdcCode"]
                teacher_params["acaAsnum"] = onehakwon["acaAsnum"]
                teacher_params["gubunCode"] = searchParams["searchGubunCode"]

                teacherList = []
                response = requests.post(neisUrl + '/hes_ica_cr91_006.ws', data=json.dumps(teacher_params), cookies=cookies)
                jsonObj = json.loads(response.content)
                teachers = jsonObj["resultSVO"]["teacherDVOList"]

                for oneteacher in teachers:
                    teacherList.append(oneteacher["fouKraName"])

            teacherTuple = tuple(teacherList)
            excelList.append(excelData + teacherTuple)
            seq = seq + 1
            progressIndex = progressIndex + 1
            printProgressBar(progressIndex, maxProgressIndex, prefix='Progress:', suffix='Complete', length=50)

        ## progress bar end
        # print("\n")

    ## 엑셀 저장
    sheet = book.active
    sheet.append(('크롤링일', '순번', '지역', '학원명', '교습과정', '교습과목', '전화번호', '주소', '정원', '교습기간', '총교습시간(분)', '교습비 합계', '교습비', '기타경비', '강사'))
    for row in excelList:
        sheet.append(row)

    filename = './data/hakwoncrawling' + '-' + areaNm + '-' + zoneNm

    if searchParams["searchClassName"] != '':
        filename = filename + '-' + searchParams["searchClassName"]

    if searchParams["searchGubunCode"] == '1':
        filename = filename + '-' + '학원'
    elif searchParams["searchGubunCode"] == '2':
        filename = filename + '-' + '교습소'

    book.save(filename + '-' + strNow + '.xlsx')

## 지역명 찾기
def findZoneName(zonecode):
    for onecode in zoneCodes:
        if onecode["zoneCode"] == zonecode:
            return onecode["zoneNm"]
    return ""

## 하위 지역 리스트 만들기
def getSearchZoneCodeList(areaIndex, cookies):
    params={}
    response = requests.post(areaUrls[areaIndex]["url"] + '/scs_ica_cr91_001.ws', data=json.dumps(params), cookies=cookies)
    jsonObj = json.loads(response.content)

    if jsonObj["result"]["status"] == "success":
        searchZoneCodeList = jsonObj["resultSVO"]["searchZoneCodeList"]
        for zonecode in searchZoneCodeList:
            zoneCodes.append({"zoneCode": zonecode["zoneCode"], "zoneNm": zonecode["zoneNm"]})
        return 1
    else:
        return 0

def readSearchConfig():
    with open('./config.json', 'rt', encoding='UTF-8') as json_file:
        p = json.load(json_file)
        searchParams["pageIndex"] = p["pageIndex"]
        searchParams["pageSize"] = p["pageSize"]
        searchParams["checkDomainCode"] = p["checkDomainCode"]
        searchParams["juOfcdcCode"] = p["juOfcdcCode"]
        searchParams["acaAsnum"] = p["acaAsnum"]
        searchParams["gubunCode"] = p["gubunCode"]
        searchParams["searchYn"] = p["searchYn"]
        searchParams["searchGubunCode"] = p["searchGubunCode"]
        searchParams["searchName"] = p["searchName"]
        searchParams["searchZoneCode"] = p["searchZoneCode"]
        searchParams["searchKindCode"] = p["searchKindCode"]
        searchParams["searchTypeCode"] = p["searchTypeCode"]
        searchParams["searchCrseCode"] = p["searchCrseCode"]
        searchParams["searchCourseCode"] = p["searchCourseCode"]
        searchParams["searchClassName"] = p["searchClassName"]


if __name__ == "__main__":
    strNow = now.strftime('%y%m%d')
    neisUrl = 'https://www.neis.go.kr'

    areaIndex = "0"
    # 검색 조건 불러오기 추가
    # config.json 에서 해당 Parameter를 미리 불러옴
    readSearchConfig()

    # 학원정보 데이터 수집 진행중
    # 2020.2.25 방순호
    # for zcode in
    while int(areaIndex) >= 0:
        ## Initialize
        searchParams["pageIndex"] = "1"
        searchParams["pageSize"] = "1"

        for aindex in range(0,len(areaUrls),1):
            print(str(aindex) + " : " + areaUrls[aindex]["areaNm"])

        areaIndex = input("교육청을 선택해 주세요 (종료 'q') : ")

        if areaIndex == 'q':
            exit(0)

        ## cookie 정보 저장
        # request = Request(neisUrl)
        params = {}
        # params["paramJson"] = "%7B%7D"
        response = requests.get(areaUrls[int(areaIndex)]["url"] + '/edusys.jsp?page=scs_m80000', params={})
        cookies = response.cookies

        ## 지역 리스트 불러오기
        zoneCodes = []
        if getSearchZoneCodeList(int(areaIndex), cookies) == 0:
            print("지역을 불러오는 도중 오류가 발생했습니다.")
            continue

        zoneIndex = ""
        while zoneIndex == "":
            for zoneCode in zoneCodes:
                print(zoneCode["zoneCode"] + " : " + zoneCode["zoneNm"])

            zoneIndex = input("원하시는 지역 번호를 입력해 주세요 (종료 'q') : ")
            if zoneIndex == 'q':
                exit(0)

            if zoneIndex.strip() != "":
                break

        searchGubun = ""
        while searchGubun == "":
            searchGubun = input("1 - 학원, 2 - 교습소 (종료 'q') : ")
            if searchGubun == 'q':
                exit(0)

            if searchGubun == '1' or searchGubun == '2':
                searchParams["searchGubunCode"] = searchGubun
                break
            else:
                searchGubun = ""

        searchWord = input("원하시는 검색어를 입력해 주세요 (없으면 걍 엔터) : ")
        searchParams["searchClassName"] = searchWord.strip()

        hakwondata(zoneIndex, areaUrls[int(areaIndex)]["url"], areaUrls[int(areaIndex)]["areaNm"], cookies)

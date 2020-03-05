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
    {'url': 'https://hakwon.gbe.go.kr', 'areaNm': '경북'},
    {'url': 'https://hakwon.gne.go.kr', 'areaNm': '경남'},
    {'url': 'https://hakwon.jje.go.kr', 'areaNm': '제주'},
]

zoneCodes = []
searchParams = {}

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

        for onehakwon in hakwonlist:
            excelData = (strNow, seq, onehakwon["zoneNm"], onehakwon["acaNm"], onehakwon["leSbjtNm"], onehakwon["faTelno"], onehakwon["totalJuso"])

            ## 강사 정보 추출
            teacher_params["juOfcdcCode"] = onehakwon["juOfcdcCode"]
            teacher_params["acaAsnum"] = onehakwon["acaAsnum"]
            teacher_params["gubunCode"] = "1"

            teacherList = []
            response = requests.post(neisUrl + '/hes_ica_cr91_006.ws', data=json.dumps(teacher_params), cookies=cookies)
            jsonObj = json.loads(response.content)
            teachers = jsonObj["resultSVO"]["teacherDVOList"]

            for oneteacher in teachers:
                teacherList.append(oneteacher["fouKraName"])

            teacherTuple = tuple(teacherList)
            excelList.append(excelData + teacherTuple)
            seq = seq + 1

    ## 엑셀 저장
    sheet = book.active
    sheet.append(('크롤링일', '순번', '지역', '학원명', '교습 과목', '전화번호', '주소', '강사'))
    for row in excelList:
        sheet.append(row)

    book.save('./data/hakwoncrawling' + '-' + areaNm + '-' + zoneNm + '-' + strNow + '.xlsx')

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

    searchZoneCodeList = jsonObj["resultSVO"]["searchZoneCodeList"]
    for zonecode in searchZoneCodeList:
        zoneCodes.append({"zoneCode": zonecode["zoneCode"], "zoneNm": zonecode["zoneNm"]})

def readSearchConfig():
    with open('./config.json', 'rt', encoding='UTF-8') as json_file:
        p = json.load(json_file)
        print(p["pageIndex"])
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
        for aindex in range(0,len(areaUrls),1):
            print(str(aindex) + " : " + areaUrls[aindex]["areaNm"])

        areaIndex = input("교육청을 선택해 주세요 (종료 'q') : ")

        if areaIndex == 'q':
            exit(0)

        ## cookie 정보 저장
        # request = Request(neisUrl)
        params = {}
        params["paramJson"] = "%7B%7D"
        response = requests.get(areaUrls[int(areaIndex)]["url"] + '/edusys.jsp?page=scs_m80000', urlencode(params).encode())
        cookies = response.cookies

        ## 지역 리스트 불러오기
        zoneCodes = []
        getSearchZoneCodeList(int(areaIndex), cookies)

        for zoneCode in zoneCodes:
            print(zoneCode["zoneCode"] + " : " + zoneCode["zoneNm"])

        zoneIndex = input("원하시는 지역 번호를 입력해 주세요 : ")

        hakwondata(zoneIndex, areaUrls[int(areaIndex)]["url"], areaUrls[int(areaIndex)]["areaNm"], cookies)

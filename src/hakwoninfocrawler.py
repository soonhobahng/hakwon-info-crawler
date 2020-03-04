from urllib.request import Request, urlopen
from urllib.parse import urlencode
from bs4 import BeautifulSoup
import csv, datetime, shutil, os, boto3, re
from botocore.exceptions import ClientError
from openpyxl import Workbook
import json
import requests

now = datetime.datetime.now()
areaUrls = [
    # 'https://hakwon.sen.go.kr/scs_ica_cr91_005.ws', # 서울
    # 'https://hakwon.pen.go.kr/scs_ica_cr91_005.ws', # 부산
    # 'https://hakwon.dge.go.kr/scs_ica_cr91_005.ws', # 대구
    # 'https://hakwon.ice.go.kr/scs_ica_cr91_005.ws', # 인천
    # 'https://hakwon.gen.go.kr/scs_ica_cr91_005.ws', # 광주
    # 'https://hakwon.dje.go.kr/scs_ica_cr91_005.ws', # 대전
    # 'https://hakwon.use.go.kr/scs_ica_cr91_005.ws', # 울산
    # 'https://hakwon.sje.go.kr/scs_ica_cr91_005.ws', # 세종시
    # 'https://hakwon.goe.go.kr/scs_ica_cr91_005.ws', # 경기도
    # 'https://hakwon.kwe.go.kr/scs_ica_cr91_005.ws', # 강원도
    # 'https://hakwon.cbe.go.kr/scs_ica_cr91_005.ws', # 충북
    # 'https://hakwon.cne.go.kr/scs_ica_cr91_005.ws', # 충남
    # 'https://hakwon.jbe.go.kr/scs_ica_cr91_005.ws', # 전북
    # 'https://hakwon.jne.go.kr/scs_ica_cr91_005.ws', # 전남
    # 'https://hakwon.gbe.go.kr/scs_ica_cr91_005.ws', # 경북
    # 'https://hakwon.gne.go.kr/scs_ica_cr91_005.ws', # 경남
    # 'https://hakwon.jje.go.kr/scs_ica_cr91_005.ws', # 제주

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

def hakwondata(zcode, neisUrl, areaNm, cookies):
    book = Workbook()
    urllist = []
    excelList = []
    teacherList = []
    params = {}
    totalPage = 1
    zoneNm = ""

    params["pageIndex"] = "1"
    params["pageSize"] = "1"
    params["checkDomainCode"] = ""
    params["juOfcdcCode"] = ""
    params["acaAsnum"] = ""
    params["gubunCode"] = ""
    params["searchYn"] = "1"
    params["searchGubunCode"] = "1"
    params["searchName"] = ""
    params["searchZoneCode"] = zcode
    params["searchKindCode"] = ""
    params["searchTypeCode"] = ""
    params["searchCrseCode"] = ""
    params["searchCourseCode"] = ""
    params["searchClassName"] = ""

    ## to get total count
    response = requests.post(neisUrl + '/scs_ica_cr91_005.ws', data=json.dumps(params), cookies=cookies)
    jsonObj = json.loads(response.content)

    ## total count
    totalCnt = int(jsonObj["resultSVO"]["totalCount"])
    print(totalCnt, "records ")

    if totalCnt > 1000:
        totalPage, therest = divmod(totalCnt, 1000)
        if therest > 0:
            totalPage = totalPage + 1
        params["pageSize"] = 1000
    else:
        totalPage = 1
        params["pageSize"] = str(totalCnt)

    print("Total ", totalPage, " pages...")

    zoneNm = findZoneName(zcode)

    for pageIndex in range(0, totalPage, 1):
        print(pageIndex + 1, " page crawling... ")
        ## post again
        params["pageIndex"] = str(pageIndex + 1)
        response = requests.post(neisUrl + '/scs_ica_cr91_005.ws', data=json.dumps(params), cookies=cookies)
        jsonObj = json.loads(response.content)
        hakwonlist = jsonObj["resultSVO"]["hesIcaCr91M00DVO"]

        seq = (pageIndex) * int(params["pageSize"]) + 1
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


if __name__ == "__main__":
    strNow = now.strftime('%y%m%d')
    neisUrl = 'https://www.neis.go.kr'

    areaIndex = "0"
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

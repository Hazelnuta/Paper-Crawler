#-*- coding:utf-8 -*- #
import requests
import threading
import time
import Queue
from time import sleep
import os
import win32com.client
import csv

def write_db_for_PC(db_num):
    ori_file = open("Paper Collection.csv")
    csv_reader = csv.reader(ori_file)
    for i in range(db_num/2 + 1):
        acc_name = 'Paper Collection' + str(i + 1)
        acc_name += '.accdb'
        if os.path.exists(acc_name):
            os.remove(acc_name)
        try:
            path = os.getcwd()
            dbname = path + '\\' + acc_name
            accApp = win32com.client.Dispatch("Access.Application")
            dbEngine = accApp.DBEngine
            workspace = dbEngine.Workspaces(0)
            dbLangGeneral = ';LANGID=0x0409;CP=1252;COUNTRY=0'
            database = workspace.CreateDatabase(dbname, dbLangGeneral, 64)
            database.Execute("""CREATE TABLE [""" + acc_name[:-7] + str(2*int(acc_name[-7]) - 1) + """](
                          PaperID autoincrement, Location varchar(45), Institution varchar(210),[Accession Number] varchar(120),
                          DOI varchar(240),[Pubmed ID] varchar(120),[Article Title] longtext,
                          Link longtext,Authors longtext,Source longtext,[Research Area] varchar(210),Volume varchar(30),
                          Issue varchar(30),Pages varchar(45),[Publication Date] varchar(24),
                          [Times Cited] varchar(45),[Journal Expected Citations] varchar(21),[Category Expected Citations] varchar(21),
                          [Journal Normalized Citation Impact] varchar(21),[Category Normalized Citation Impact] varchar(30),
                          [Percentile in Subject Area] varchar(15),[Journal Impact Factor] varchar(21) );""")
            database.Execute("""CREATE TABLE [""" + acc_name[:-7] + str(2*int(acc_name[-7])) + """](
                          PaperID autoincrement, Location varchar(45), Institution varchar(210),[Accession Number] varchar(120),
                          DOI varchar(240),[Pubmed ID] varchar(120),[Article Title] longtext,
                          Link longtext,Authors longtext,Source longtext,[Research Area] varchar(210),Volume varchar(30),
                          Issue varchar(30),Pages varchar(45),[Publication Date] varchar(24),
                          [Times Cited] varchar(45),[Journal Expected Citations] varchar(21),[Category Expected Citations] varchar(21),
                          [Journal Normalized Citation Impact] varchar(21),[Category Normalized Citation Impact] varchar(30),
                          [Percentile in Subject Area] varchar(15),[Journal Impact Factor] varchar(21) );""")
        except Exception as e:
            print(e)
        finally:
            accApp.DoCmd.CloseDatabase
            accApp.Quit
            database = None
            workspace = None
            dbEngine = None
            accApp = None
        conn = win32com.client.Dispatch(r"ADODB.Connection")
    flag1 = -1
    for line in csv_reader:
        flag1 += 1
        if flag1 % 1800000 == 0:
            try:
                conn.close()
            except:
                pass
            DSN = 'PROVIDER = Microsoft.Jet.OLEDB.4.0;DATA SOURCE = Paper Collection' + str(
                flag1 / 1800000 + 1) + ".accdb"
            conn.Open(DSN)
        if flag1 == 0:
            continue
        flag2 = 0
        sql = "INSERT INTO [Paper Collection" + str(flag1 / 900000 + 1) + "]" \
               "(Location, Institution,[Accession Number]," \
               "DOI,[Pubmed ID],[Article Title]," \
               "Link,Authors,Source,[Research Area],Volume," \
               "Issue,Pages,[Publication Date]," \
               "[Times Cited],[Journal Expected Citations],[Category Expected Citations]," \
               "[Journal Normalized Citation Impact],[Category Normalized Citation Impact]," \
               " [Percentile in Subject Area],[Journal Impact Factor])VALUES ("
        for item in line:
            flag2 += 1
            if flag2 == 1:
                continue
            if flag2 == 22:
                item = item[0:len(item) - 1]
            if item.find('wosUrl')!= -1:
                item = item[:-8]
                while item[-1] == '\"':
                    item = item[:-1]
            item = '\"' + item
            item = item + '\"'
            sql = sql + item
            sql = sql + ','
            #print item
        sql = sql[0:len(sql) - 2]
        sql = sql + "\")"
        try:
            conn.execute(sql)
        except Exception as e:
            print(e)
            print line[6]
            pass
            # conn.commit()
    ori_file.close()

def write_db_for_RA():
    if os.path.exists('Research Area Info Collection.accdb'):
        os.remove('Research Area Info Collection.accdb')
    try:
        dbname = path + r'\Research Area Info Collection.accdb'
        accApp = win32com.client.Dispatch("Access.Application")
        dbEngine = accApp.DBEngine
        workspace = dbEngine.Workspaces(0)
        dbLangGeneral = ';LANGID=0x0409;CP=1252;COUNTRY=0'
        database = workspace.CreateDatabase(dbname, dbLangGeneral, 64)
        database.Execute("""CREATE TABLE [Research Area Info Collection] (
                                  ID autoincrement,
                                  Institution varchar(70),
                                  [Research Area] varchar(30),
                                  Rank varchar(10),
                                  [Web of Science Documents] varchar(10),
                                  [Category Normalized Citation Impact] varchar(10),
                                  [Times Cited] varchar(10),
                                  [% Docs Cited] varchar(10),
                                  [% Documents in Q1 Journals] varchar(10),
                                  [% Documents in Q2 Journals] varchar(10),
                                  [% Documents in Q3 Journals] varchar(10),
                                  [% Documents in Q4 Journals] varchar(10),
                                  [% Documents in Top 1%] varchar(10),
                                  [% Documents in Top 10%] varchar(10),
                                  [% Highly Cited Papers] varchar(10),
                                  [% Hot Papers] varchar(10),
                                  [% Industry Collaborations] varchar(10),
                                  [% International Collaborations] varchar(10),
                                  [Average Percentile] varchar(10),
                                  [Citation Impact] varchar(10),
                                  [Documents in JIF Journals] varchar(10),
                                  [Documents in Q1 Journals] varchar(10),
                                  [Documents in Q2 Journals] varchar(10),
                                  [Documents in Q3 Journals] varchar(10),
                                  [Documents in Q4 Journals] varchar(10),
                                  [Highly Cited Papers] varchar(10),
                                  [Impact Relative to World] varchar(10),
                                  [International Collaborations] varchar(10),
                                  [Journal Normalized Citation Impact] varchar(10)
                                  );""")
    except Exception as e:
        print(e)
    finally:
        accApp.DoCmd.CloseDatabase
        accApp.Quit
        database = None
        workspace = None
        dbEngine = None
        accApp = None
    conn = win32com.client.Dispatch(r"ADODB.Connection")
    DSN = 'PROVIDER = Microsoft.Jet.OLEDB.4.0;DATA SOURCE = Research Area Info Collection.accdb'
    conn.Open(DSN)
    ori_file = open("Research Area Info Collection.csv")
    csv_reader = csv.reader(ori_file)
    flag1 = 0
    for line in csv_reader:
        if flag1 == 0:
            flag1 = 1
            continue
        flag2 = 0
        sql = "INSERT INTO [Research Area Info Collection]" \
              "(Institution,[Research Area],Rank,[Web of Science Documents],[Category Normalized Citation Impact]," \
              "[Times Cited],[% Docs Cited],[% Documents in Q1 Journals],[% Documents in Q2 Journals]," \
              "[% Documents in Q3 Journals],[% Documents in Q4 Journals],[% Documents in Top 1%],[% Documents in Top 10%]," \
              "[% Highly Cited Papers],[% Hot Papers],[% Industry Collaborations],[% International Collaborations]," \
              "[Average Percentile],[Citation Impact] ,[Documents in JIF Journals],[Documents in Q1 Journals]," \
              "[Documents in Q2 Journals],[Documents in Q3 Journals],[Documents in Q4 Journals],[Highly Cited Papers]," \
              "[Impact Relative to World],[International Collaborations],[Journal Normalized Citation Impact])VALUES ("
        for item in line:
            flag2 += 1
            if flag2 == 1:
                continue
            if flag2 == 29:
                item = item[0:len(item) - 1]
            item = '\"' + item
            item = item + '\"'
            sql = sql + item
            sql = sql + ','
        sql = sql[0:len(sql) - 2]
        sql = sql + "\")"
        # print sql
        try:
            conn.execute(sql)
        except:
            pass
    ori_file.close()

def get_info_by_RA(name, number):
    url = "https://incites.thomsonreuters.com/incites-app/explore/0/subject/data/table/page/export/csv"
    querystring = {"fileName": "InCites Research Areas"}
    #payload = "{\"take\":10000000,\"skip\":0,\"sortBy\":\"timesCited\",\"sortOrder\":\"desc\",\"filters\":{\"location\":{\"is\":[\"CHINA MAINLAND\",\"TAIWAN\",\"MACAU\",\"HONG KONG\"]},\"personIdTypeGroup\":{\"is\":\"name\"},\"personIdType\":{\"is\":\"fullName\"},\"schema\":{\"is\":\"Essential Science Indicators\"},\"articletype\":{\"is\":[\"Article\",\"Letter\",\"Review\"]},\"period\":{\"is\":[2007,2017]}},\"pinned\":[],\"indicators\":[\"orgName\",\"rank\",\"percentCited\",\"prcntDocsIn99\",\"prcntDocsIn90\",\"prcntHighlyCitedPapers\",\"prcntHotPapers\",\"prcntIndCollab\",\"prcntIntCollab\",\"acadStaffStdnt\",\"acadStaffInt\",\"avrgPrcnt\",\"norm\",\"ncicountry\",\"avrgCitations\",\"location\",\"doctoral\",\"doctoralUndergrad\",\"docsCited\",\"esi\",\"hindex\",\"highlyCitedPapers\",\"impactRelToWorld\",\"instIncome\",\"intCollaborations\",\"jNCI\",\"level\",\"type\",\"papers\",\"papersInt\",\"resIncome\",\"resIncomeInd\",\"resReputGlob\",\"stateProvice\",\"stdntInt\",\"teachingReput\",\"the\",\"timesCited\",\"wosDocuments\",\"key\",\"seqNumber\",\"hasProfile\"]}"
    # curPayload2 = "params={\"filters\":{\"schema\":{\"is\":\"Essential Science Indicators\"},\"assprsnIdTypeGroup\":{\"is\":\"name\"},\"assprsnIdType\":{\"is\":\"fullName\"},\"orgname\":{\"is\":[\"Chinese Academy of Sciences\"]},\"location\":{\"is\":[\"CHINA MAINLAND\",\"TAIWAN\",\"HONG KONG\",\"MACAU\"]},\"articletype\":{\"is\":[\"Article\",\"Letter\",\"Review\"]},\"personIdTypeGroup\":{\"is\":\"name\"},\"personIdType\":{\"is\":\"fullName\"},\"period\":{\"is\":[2014,2017]}},\"skip\":0,\"take\":22,\"sortBy\":\"timesCited\",\"sortOrder\":\"desc\",\"pinned\":[],\"indicators\":[\"sbjName\",\"rank\",\"wosDocuments\",\"norm\",\"timesCited\",\"percentCited\",\"percjifdocsq1\",\"percjifdocsq2\",\"percjifdocsq3\",\"percjifdocsq4\",\"prcntDocsIn99\",\"prcntDocsIn90\",\"prcntHighlyCitedPapers\",\"prcntHotPapers\",\"prcntIndCollab\",\"prcntIntCollab\",\"avrgPrcnt\",\"avrgCitations\",\"jifdocs\",\"jifdocsq1\",\"jifdocsq2\",\"jifdocsq3\",\"jifdocsq4\",\"highlyCitedPapers\",\"impactRelToWorld\",\"intCollaborations\",\"jNCI\",\"key\",\"seqNumber\"],\"benchmarkNames\":[],\"dateInfo\":{}}"
    flag = payload.find("location")
    flag1 = payload.find("pinned")
    mid = payload[flag - 1:flag1 - 2]
    flag = name.find("&")
    if flag != -1:
        newname = name[:flag] + "%26" + name[flag+1:]
        name = newname
    curPayload2 = "params={\"filters\":{\"assprsnIdTypeGroup\":{\"is\":\"name\"},\"assprsnIdType\":{\"is\":\"fullName\"},\"orgname\":{\"is\":[\""+name+"\"]}," \
                  + mid + ",\"skip\":0,\"take\":100,\"sortBy\":\"timesCited\",\"sortOrder\":\"desc\",\"pinned\":[],\"indicators\":[\"sbjName\",\"rank\",\"wosDocuments\",\"norm\",\"timesCited\",\"percentCited\",\"percjifdocsq1\",\"percjifdocsq2\",\"percjifdocsq3\",\"percjifdocsq4\",\"prcntDocsIn99\",\"prcntDocsIn90\",\"prcntHighlyCitedPapers\",\"prcntHotPapers\",\"prcntIndCollab\",\"prcntIntCollab\",\"avrgPrcnt\",\"avrgCitations\",\"jifdocs\",\"jifdocsq1\",\"jifdocsq2\",\"jifdocsq3\",\"jifdocsq4\",\"highlyCitedPapers\",\"impactRelToWorld\",\"intCollaborations\",\"jNCI\",\"key\",\"seqNumber\"],\"benchmarkNames\":[],\"dateInfo\":{}}"
    headers = {
        'origin': "https://incites.thomsonreuters.com",
        'upgrade-insecure-requests': "1",
        'content-type': "application/x-www-form-urlencoded",
        'user-agent': "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.139 Safari/537.36",
        'accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
        'referer': "https://incites.thomsonreuters.com/",
        'accept-encoding': "gzip, deflate, br",
        'accept-language': "en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7",
        'cookie': "JSESSIONID=B174A778AF4E5EA0955D39CF8AECF025; _ga=GA1.2.1626134384.1526014950; _gid=GA1.2.314431314.1526014950; USERNAME=\"hazelnut@sjtu.edu.cn\"; PSSID=\"H3-b8er91Ez5x2FwCGnGatiRavAqIaHLPzALd-18x2dHZc1BbvQhu1TPTpUU0DGKQx3Dx3DpM4NQChEy0Vl6BB5Y3ZA0gx3Dx3D-9vvmzcndpRgQCGPd1c2qPQx3Dx3D-wx2BJQh9GKVmtdJw3700KssQx3Dx3D\"; CUSTOMER_NAME=\"SHANGHAI JIAO TONG UNIV\"; E_GROUP_NAME=\"IC2 Platform\"; SUBSCRIPTION_GROUP_ID=\"242426\"; SUBSCRIPTION_GROUP_NAME=\"SHANGHAI JIAO TONG UNIVERSITY_Benchmarking\"; CUSTOMER_GROUP_ID=\"93841\"; ROAMING_DISABLED=\"false\"; ACCESS_METHOD=\"UNP\"; SECCONTEXT=34b44ec82e72d56ddd6b588c3388abe41a41c42d16cdeebd212d6ca92435bc44b19f8fda89856e2477df0910a99e12cf96ce137a843879ef2d6b8b3e8c69f8a4b4ca70154744cd0303f78268fadbe6b6513363ed95542902eab9ed6dc812dedc1c9654dc3482081864eb64a65e13d61a165e71b8639831f550e7efbc29d517e1c73fda6b3613b8d07a79a1753aa26993009baf44f96401cf0f3ebf7168bf8190c71c55d3c9adc6f30b93ad545f707e4b0fece282f4213a5c2d2d15218914ebc9a22a123c5dd08cb94d821a084ba6d940d36d973f03d5f4eb558eb73c5d16a5916948a034d9a971b6d709e8c68f48fae4b0b6f057f524c7bac515c0d2d3c2bc6a6082bf8347d920b58b95f87ad47666d35d01aac9439cb1c820377143f043f9dec53b68c6f63f822eaba06bb2dc9090c0961d4ad53111b36800900804e8b227f3654f7ddbc2344bca0d7463bff01d12159c55c2063a39e4f55b4c980d2f56b2da5dcd45f9cbaf2becd06122b5fe933b08132053bf588726ca0ef125b0a7e5245328d0abeccdbb069aaddb2151e24fc01f2ab799e6f204410e8f878b36a99dbc6d5870c9b54a7fbec40981b1269572158a3886b25e0f5c6d7036d269029ce46af68501f1c69be3028c391a52d46d71eab6aba4788f9a8932688527eb7a7b677fb016618f815960831f03b514600cdbb71d2baaa7c7655bb4484782bc25323b43b28aff0dfbe664f2cc174bca68c25e0e25; _gat=1",
        'cache-control': "no-cache",
        'postman-token': "abbb3218-e2b6-278c-f2fd-fb750fc53ccb"
    }
    if threading.active_count() > 5:
        while number % 49 != int(100 * time.time()) % 49:
            sleep(1)
    response = requests.request("POST", url, data=curPayload2, headers=headers, params=querystring)
    content = response.text
    ra_list = content.split("\n")
    ra_list = ra_list[1:-12]
    return ra_list

def get_RA(name, number):
    ra_list = get_info_by_RA(name, number)
    filename = curPath + '\\' + name + '.csv'
    myfile = open(filename, 'wb')
    csv_writer = csv.writer(myfile, dialect="excel")
    TitleLine = ["Institution","Research Area","Rank","Web of Science Documents","Category Normalized Citation Impact","Times Cited",
                 "% Docs Cited","% Documents in Q1 Journals","% Documents in Q2 Journals",
                 "% Documents in Q3 Journals","% Documents in Q4 Journals","% Documents in Top 1%",
                 "% Documents in Top 10%","% Highly Cited Papers","% Hot Papers","% Industry Collaborations",
                 "% International Collaborations","Average Percentile","Citation Impact",
                 "Documents in JIF Journals","Documents in Q1 Journals","Documents in Q2 Journals",
                 "Documents in Q3 Journals","Documents in Q4 Journals","Highly Cited Papers",
                 "Impact Relative to World","International Collaborations","Journal Normalized Citation Impact"]
    csv_writer.writerow(TitleLine)
    print('writing ' + name + '.csv')

    for item in ra_list:
        if item[0] == "\"":
            item = item[1:]
            item = item.split("\"")
            item[1] = item[1].split(",")
            item[1] = item[1][1:]
            item = [item[0]] + item[1]
        else:
            item = item.split(",")
        item = [name] + item
        csv_writer.writerow(item)
    myfile.close()
    print(name + '.csv written already')

def getMoreinfo(key, num, skip, name, number, year, order):
    url = "https://incites.thomsonreuters.com/incites-app/drilldowns/0/organization/dbd_39/data/export/csv"
    #querystring = {"skip": skip, "sortBy": "cites", "sortOrder": "desc", "take": num,"fileName": name, "key":key}
    if order == 1:
        querystring = {"skip": skip, "sortBy": "cites", "sortOrder": "asc", "take": "10000", "fileName": name,
                       "key": key}
    else:
        querystring = {"skip": skip, "sortBy": "cites", "sortOrder": "desc", "take": "10000", "fileName": name,
                       "key": key}
    # payload = "{\"take\":10000000,\"skip\":0,\"sortBy\":\"timesCited\",\"sortOrder\":\"desc\",\"filters\":{\"location\":{\"is\":[\"CHINA MAINLAND\",\"TAIWAN\",\"MACAU\",\"HONG KONG\"]},\"personIdTypeGroup\":{\"is\":\"name\"},\"personIdType\":{\"is\":\"fullName\"},\"schema\":{\"is\":\"Essential Science Indicators\"},\"articletype\":{\"is\":[\"Article\",\"Letter\",\"Review\"]},\"period\":{\"is\":[2007,2017]}},\"pinned\":[],\"indicators\":[\"orgName\",\"rank\",\"percentCited\",\"prcntDocsIn99\",\"prcntDocsIn90\",\"prcntHighlyCitedPapers\",\"prcntHotPapers\",\"prcntIndCollab\",\"prcntIntCollab\",\"acadStaffStdnt\",\"acadStaffInt\",\"avrgPrcnt\",\"norm\",\"ncicountry\",\"avrgCitations\",\"location\",\"doctoral\",\"doctoralUndergrad\",\"docsCited\",\"esi\",\"hindex\",\"highlyCitedPapers\",\"impactRelToWorld\",\"instIncome\",\"intCollaborations\",\"jNCI\",\"level\",\"type\",\"papers\",\"papersInt\",\"resIncome\",\"resIncomeInd\",\"resReputGlob\",\"stateProvice\",\"stdntInt\",\"teachingReput\",\"the\",\"timesCited\",\"wosDocuments\",\"key\",\"seqNumber\",\"hasProfile\"]}"
    #curPayload1 = "params={\"filters\":{\"location\":{\"is\":[\"CHINA MAINLAND\",\"TAIWAN\",\"HONG KONG\",\"MACAU\"]},\"personIdTypeGroup\":{\"is\":\"name\"},\"personIdType\":{\"is\":\"fullName\"},\"schema\":{\"is\":\"Essential Science Indicators\"},\"articletype\":{\"is\":[\"Article\",\"Review\",\"Letter\"]},\"period\":{\"is\":[2007,2017]}},\"pinned\":[],\"dateInfo\":{\"exportDate\":\"2017-08-06\",\"wosDate\":\"2017-05-31\",\"deployDate\":\"2017-07-22\"}}"
    flag = payload.find("\"filters\"")
    flag1 = payload.find("pinned")
    if year != 0:
        flag2 = payload.find('20')
        payload1 = payload[:flag2] + str(year) + ',' + str(year + 1) + payload[flag2 + 9:]
    else:
        payload1 = payload
    curPayload = "params={" + payload1[flag: flag1 + 11] + "\"dateInfo\":{\"exportDate\":\"2017-08-06\",\"wosDate\":\"2017-05-31\",\"deployDate\":\"2017-07-22\"}}"
    headers = {
        'origin': "https://incites.thomsonreuters.com",
        'upgrade-insecure-requests': "1",
        'content-type': "application/x-www-form-urlencoded",
        'user-agent': "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.181 Safari/537.36",
        'accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
        'referer': "https://incites.thomsonreuters.com/",
        'accept-encoding': "gzip, deflate, br",
        'accept-language': "en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7",
        'cookie': "_ga=GA1.2.1988740908.1527559100; _gid=GA1.2.501115603.1527559100; USERNAME=\"hazelnut@sjtu.edu.cn\"; CUSTOMER_NAME=\"SHANGHAI JIAO TONG UNIV\"; E_GROUP_NAME=\"IC2 Platform\"; SUBSCRIPTION_GROUP_ID=\"242426\"; SUBSCRIPTION_GROUP_NAME=\"SHANGHAI JIAO TONG UNIVERSITY_Benchmarking\"; CUSTOMER_GROUP_ID=\"93841\"; ROAMING_DISABLED=\"false\"; ACCESS_METHOD=\"UNP\"; _gat=1; JSESSIONID=CCCA0D0F2C717AE594199E625E609A41; PSSID=\"H1-D90zu5x2BmPnPHjHjHvjaqVeYZmjO3aaIm-18x2d4Z4Z6kWVnx2B4PrQKQ8eH2vwx3Dx3DyyZOoXUx2Fgq0uvKP57JBruQx3Dx3D-iyiHxxh55B2RtQWBj2LEuawx3Dx3D-1iOubBm4x2FSwJjjKtx2F7lAaQx3Dx3D\"; SECCONTEXT=a4e127693e537ca470fa3e0ca0932398b8ff9449dbef958eef72dc7b3e216f14188ba490ba80ccef20a295c4b4f072f8079c62c19d2bd961d0fc7ae7dbbb717d0bf5a12b86d8bcc7967ac9ae9b2151dea64ff8dec8b2bad7b005178896bb1fa2d28fb106d0f8b5e80b08c5e89e329a22697db49e7326dae37a3599b6dbf5292b15386d0daf34dc879dc30a4d5cca698e6d17ca9040456bc316a8cb1f0540fdf65b1c56f90e37f07ceff37948e71d89e09ffc153fa152d01172b362de9926463cc73c88d6d230bb3a9100b406fa87115772ed5cc901422a45148e61dc78312ca062da4f1f4cdbd11bfb5f2d6b3baaacf13fb6b99347b03e389afe5edf02c64e2389f07bed70f14227fe6d2c7b725b1e7326e0d64ea5c11089a1f47ca1624de53308f88c25310c3aeb932a4222a6b0690dee6176c77e2e031d82a960f74aa3184e32ad877d18f857450e46121206c24fbd5b01e8f75597f643694825c760ff5532711abf2925fdcfa41416484408bc4f70cce1bf5a050dabcad23c9661edbc16878ce7aa3a97cad0706392d4958a5560aa603d845657c3a19f83ff187d238d5ee9b65856884bde4c43d7b35419e6c55bd1e05c4ca9e2f84819ed72c80888978543f584088be6232c7c87671138020da1ccf9755008df263d7b3e573f61f3e0525c64c27812bd3621ab6a28da05466c45cce86d2cfeb7d456558335b82702af7111849529565e61d71ae16237402bcb8093; JSESSIONID=CB9C33241E20B71A0D3E9C7258989EF3; _ga=GA1.2.1988740908.1527559100; CUSTOMER_NAME=\"SHANGHAI JIAO TONG UNIV\"; ROAMING_DISABLED=\"false\"; _gid=GA1.2.242436379.1527747230; USERNAME=\"hazelnut@sjtu.edu.cn\"; PSSID=\"H2-2NqOi4YC6eL5cOB9VqufdMazkwod2Qnx2B-18x2dUE4r2u4PqmoPrQKQ8eH2vwx3Dx3DKh3mcv4TPYtO7vsgn6Jz8wx3Dx3D-iyiHxxh55B2RtQWBj2LEuawx3Dx3D-1iOubBm4x2FSwJjjKtx2F7lAaQx3Dx3D\"; E_GROUP_NAME=\"IC2 Platform\"; SUBSCRIPTION_GROUP_ID=\"242426\"; SUBSCRIPTION_GROUP_NAME=\"SHANGHAI JIAO TONG UNIVERSITY_Benchmarking\"; CUSTOMER_GROUP_ID=\"93841\"; ACCESS_METHOD=\"IP_ROAMING\"; SECCONTEXT=cc671d6c5028f4334aa537f7939cb6cda8548a0ba88db1d853522d889104a8bdd037b66fadb65699cdc09ad5721df7709148583b75fbde730a838ab59b7e9fcbe91a90c9eafadae11896f4318caafd54fbcf685bc02172879dd6dbec3273ed5c6ed8b6891f75e32768ac25ccbd6a13e24b76888e276c14c8b1576657543b0b5d935c3e1df429032ab983e071d69dcd930a1e5f8d37d0942d683467bd60382e255a486fdfaa26aae87146efbbb9a91700a7a0cd3418c9f7aed3ad1fc4df6a3f06956a1ce3c38d780d2deccd4c06a1a004e85675e73fa43ff542eedd77a6937d0c21a6d9019774d95a10f4a9f96fd5686e8135f0c780df6cd0e03c3dc6a0943864a2046ff87b6b18fcfbf44e2ffdf3729c0947113ccb94f98afe1f998978d0875b803792bd23e0cd0a69a05d107f247d41163f125574e3c84531808be16429e9c7e9dadc12f3b25f68f97234d9a6275c3f33fa82d56fa8b1ff006e4f55966c2320ddec24a50fee528cc25e17165becd27e28528a089ae79804e737f50579972afbdd5f5fd033ae1136193504d1b4bfc7f2dc746a83cf9d7ea2ccaa720bcf6d279a15fb01153d4ac579268ee3ab8d5016749a929e553b7ea6a5c1accc91007bb68c3d466e0137816757503543e5dd931a9dcd2afb44f2f22906e106bd8db8232b838698102c067fad75e90dc35e390827015d8cdb7363341d4db9b1f57e20e5fd8d722d17628376a07e22b5e9fb1c091a62; _gat=1",
        'cache-control': "no-cache",
        'postman-token': "d67865fe-7806-89a4-2c73-f7e602db5fe3"
    }
    if threading.active_count() > 5:
        while number%49 != int(100*time.time())%49:
            sleep(float(float(3)/50*threading.active_count()))
    else:
        sleep(1)
    response = requests.request("POST", url, data=curPayload, headers=headers, params=querystring)
    content = response.text
    paperList = content.split('\n')
    paperList = paperList[1:]
    while len(paperList) > 0 and paperList[-1].find("WOS") == -1:
        paperList = paperList[:-1]
    if len(paperList) == 1 and paperList[0].find("WOS") == -1:
        print content
        return []
    if len(paperList) == 0:
        print name + ": got nothing this time"
        sleep(30)
        #print content
    print name+" get more papers: " + str(len(paperList))
    return paperList

def getRow(paper, location, name):
    paperRow = [location, name]
    flag = paper.find("http")
    itemList1 = paper[:flag - 1]
    itemList2 = paper[flag:]
    itemList1 = itemList1.split(",")
    part = itemList1[3:]
    part = ",".join(part)
    while part[0] == "\"":
        part = part[1:]
    while part[-1] == "\"":
        part = part[0:-1]
    tmp = []
    tmp = tmp + itemList1[0:3] + [part]
    itemList1 = tmp

    itemList2 = itemList2.split("\"")
    partNum = len(itemList2)
    tmp = []
    for i in range((partNum - 1) / 2):
        part = itemList2[2 * i]
        part = part.split(",")
        if part == ["", ""]:
            tmp += [itemList2[2 * i + 1]]
            continue
        while part[0] == "":
            part = part[1:]
        while part[-1] == "":
            part = part[0:-1]
        tmp += part
        tmp += [itemList2[2 * i + 1]]
    part = itemList2[-1].split(",")
    while part[0] == "":
        part = part[1:]
    while part[-1] == "":
        part = part[0:-1]
    tmp += part
    paperRow = paperRow + itemList1 + tmp
    return paperRow

def getInfo(key, num, name, location, number):
    print "now we're gonna get " + str(num) +" papers from " + name + ":)"
    paperList = []
    if num >= 50000:
        print "hard mode starts for " + name
        for year in range(int(startYear), int(endYear)+1):
            print name + ": "+str(year) + '-' + str(year + 1)
            skip = 0
            zero_count = 0
            order = 0
            this_time_count = 0
            while True:
                if len(paperList) >= num - 5:
                    break
                #print str(this_time_count)
                if this_time_count >= 49990:
                    print name + " ("+ str(year) + ", " + str(year + 1) + ")"+": reverse now:)"
                    order = 1
                    skip = 0
                    this_time_count = 0
                new_papers = getMoreinfo(key, num, skip, name, number, year, order)
                skip += 10000
                this_time_count += len(new_papers)
                original_lenth = len(paperList)
                for new_item in new_papers:
                    if new_item not in paperList:
                        paperList.append(new_item)
                if original_lenth == len(paperList) and len(new_papers) != 0:
                    break
                if len(new_papers) < 9999 and len(new_papers) > 0:
                    print (name + ' currently gets '+ str(len(paperList))+ ' papers')
                    break
                if len(new_papers) == 0:
                    zero_count = zero_count + 1
                    if zero_count == 3:
                        break
    else:
        times = int(num)/10000
        for tm in range(times + 1):
            skip = tm*10000
            skip = str(skip)
            paperList += getMoreinfo(key, num, skip, name, number,0, 0)
        if len(paperList) < num and len(paperList) > 0:
            print name + ": seems like some papers have been lost and we have to do it again:("
            getInfo(key, num, name, location, number)
            return
    print (name + " current papers: " + str(len(paperList) - 1))
    filename = curPath + '\\' + name + '.csv'
    myfile = open(filename, 'wb')
    csv_writer = csv.writer(myfile, dialect="excel")
    TitleLine = ["PaperID", "Location", "Institution", "Accession Number", "DOI", "Pubmed ID", "Article Title", "Link",
                 "Authors", "Source", "Research Area", "Volume", "Issue",
                 "Pages", "Publication Date", "Times Cited", "Journal Expected Citations",
                 "Category Expected Citations",
                 "Journal Normalized Citation Impact", "Category Normalized Citation Impact",
                 "Percentile in Subject Area",
                 "Journal Impact Factor"]
    csv_writer.writerow(TitleLine)
    print ('writing ' + name + '.csv')
    paperid = 1
    for paper in paperList:
        paperRow = getRow(paper, location, name)
        paperRow = [paperid] + paperRow
        if paperRow[7].find('gateway') == -1:
            paperRow[6] = paperRow[6] + '(' + paperRow[7] + ')'
            paperRow = paperRow[0:7] + paperRow[8:]
        if paperRow[5][-1] == "\"":
            paperRow[4] = paperRow[4] + paperRow[5]
            paperRow[4] = paperRow[4][1:-1]
            comma = paperRow[6].find(",")
            paperRow[5] = paperRow[6][:comma]
            paperRow[6] = paperRow[6][comma + 1:]
        if paperRow[6][0] == "\"":
            paperRow[6] = paperRow[6][1:]
        if len(paperRow) == 21:
            paperRow = paperRow[:8] + ['n/a'] + paperRow[8:]
        csv_writer.writerow(paperRow)
        paperid += 1
    myfile.close()
    print (name + '.csv written already')


def work(number):
    while not q.empty():
        print "--------------current active threading number: " + str(threading.active_count()) + "---------------"
        #if varLock.acquire():
        k = q.get()
            #varLock.release()
        item = itemList[k]
        #print item
        name = key = num = location = ""
        if option == 1:
            flag = item.find('key=')
            key = item[flag + 4:flag + 18]
            if key[0] > '9' or key[0] < '0':
                return
            flag = item.find('value')
            flag2 = item[flag + 7:flag + 17].find(',')
            num = item[flag + 7:flag + 6 + flag2]
            flag = item.find('orgName')
            flag2 = item[flag + 10:flag + 80].find(',')
            name = item[flag + 10:flag + 10 + flag2 - 1]
            flag = item.find('location')
            flag2 = item[flag + 11:flag + 50].find('esi')
            location = item[flag + 11:flag + 11 + flag2 - 3]
            print key
            print num
            print name
            print location
        else:
            flag = item.find('orgName')
            flag2 = item[flag + 10:flag + 80].find(',')
            name = item[flag + 10:flag + 10 + flag2 - 1]
        fn = curPath + '\\' + name + '.csv'
        if os.path.exists(fn):
            print name + ' file existed:)'
            if len(os.listdir(curPath)) == TOTAL:
                return
            continue
        else:
            if option == 1:
                getInfo(key, int(num), name, location, number)
            else:
                get_RA(name, number)
            print (str(k + 1) + ' institutions already done')
        print '*********************'
        if len(os.listdir(curPath)) == TOTAL:
            return

#############################################################################################################################################50
NUM = input("How many threads do you want (50 recommended): ")
option = 0
while option != 1 and option != 2:
    option = raw_input("Do you want \nA. detailed paper info of all institutions\nB. research area info of all institutions\n[A/B]:")
    if option == "A":
        option = 1
    elif option == "B":
        option = 2
    else:
        option = input("Illegal input. Please try again [A\B]: ")

url = "https://incites.thomsonreuters.com/incites-app/explore/0/organization/data/table/page"

path = os.getcwd()
if option == 1:
    curPath = path + '\PAPERS'
else:
    curPath = path + '\Research Area'
if not os.path.exists(curPath):
    os.mkdir(curPath)

myfile = open("config.txt", 'r')
list = myfile.readlines()
myfile.close()

payload = "{\"take\":10000,\"skip\":0,\"sortBy\":\"timesCited\",\"sortOrder\":\"desc\",\"filters\":{"

startYear = list[0][-6:-2]
endYear = list[1][-6:-2]
print startYear
print endYear

if list[2].find("(") == -1:
    locRestr = 0
    print "no restriction on location."
else:
    locRestr = 1
    location = list[2].split('(')
    location = location[1:]
    for i in range(len(location)):
        location[i] = location[i][:-1]
        if location[i][-1] == ')':
            location[i] = location[i][:-1]
        location[i] = "\"" + location[i] + "\""
        print location[i]
    payload += "\"location\":{\"is\":["
    for loc in location:
        payload += loc
        payload += ','
    payload = payload[:-1]
    payload += "]},"
print '******'

payload += "\"personIdTypeGroup\":{\"is\":\"name\"},\"personIdType\":{\"is\":\"fullName\"}"

if list[3].find("(") == -1:
    schRestr = 0
    print "no restriction on schema."
else:
    schRestr = 1
    schema = list[3].split('(')
    schema = schema[1:]
    for i in range(len(schema)):
        schema[i] = schema[i][:-1]
        if schema[i][-1] == ')':
            schema[i] = schema[i][:-1]
        schema[i] = "\"" + schema[i]+ "\""
        print schema[i]
    payload += ",\"schema\":{\"is\":"
    for schm in schema:
        payload += schm
        payload += ','
    payload = payload[:-1]
    payload += "}"
print '******'

if list[4].find("(") == -1:
    artRestr = 0
    print "no restriction on article type."
else:
    artRestr = 1
    artiType = list[4].split('(')
    artiType = artiType[1:]
    for i in range(len(artiType)):
        artiType[i] = artiType[i][:-1]
        if artiType[i][-1] == ')':
            artiType[i] = artiType[i][:-1]
        artiType[i] = "\"" + artiType[i] + "\""
        print artiType[i]
    payload += ",\"articletype\":{\"is\":["
    for artp in artiType:
        payload += artp
        payload += ','
    payload = payload[:-1]
    payload += "]}"
print '******'

payload += ",\"period\":{\"is\":["
payload += startYear
payload += ','
payload += endYear
payload += "]}},\"pinned\":[],\"indicators\":[\"orgName\",\"rank\",\"wosDocuments\",\"norm\",\"timesCited\",\"percentCited\",\"percjifdocsq1\",\"percjifdocsq2\",\"percjifdocsq3\",\"percjifdocsq4\",\"prcntDocsIn99\",\"prcntDocsIn90\",\"prcntHighlyCitedPapers\",\"prcntHotPapers\",\"prcntIndCollab\",\"prcntIntCollab\",\"acadStaffStdnt\",\"acadStaffInt\",\"avrgPrcnt\",\"ncicountry\",\"avrgCitations\",\"location\",\"doctoral\",\"doctoralUndergrad\",\"docsCited\",\"jifdocs\",\"jifdocsq1\",\"jifdocsq2\",\"jifdocsq3\",\"jifdocsq4\",\"esi\",\"hindex\",\"highlyCitedPapers\",\"impactRelToWorld\",\"instIncome\",\"intCollaborations\",\"jNCI\",\"level\",\"type\",\"papers\",\"papersInt\",\"resIncome\",\"resIncomeInd\",\"resReputGlob\",\"stateProvice\",\"stdntInt\",\"teachingReput\",\"key\",\"seqNumber\",\"hasProfile\"]}"
# payload = "{\"take\":10000000,\"skip\":0,\"sortBy\":\"timesCited\",\"sortOrder\":\"desc\",\"filters\":{\"location\":{\"is\":[\"CHINA MAINLAND\",\"TAIWAN\",\"MACAU\",\"HONG KONG\"]},\"personIdTypeGroup\":{\"is\":\"name\"},\"personIdType\":{\"is\":\"fullName\"},\"schema\":{\"is\":\"Essential Science Indicators\"},\"articletype\":{\"is\":[\"Article\",\"Letter\",\"Review\"]},\"period\":{\"is\":[2007,2017]}},\"pinned\":[],\"indicators\":[\"orgName\",\"rank\",\"percentCited\",\"prcntDocsIn99\",\"prcntDocsIn90\",\"prcntHighlyCitedPapers\",\"prcntHotPapers\",\"prcntIndCollab\",\"prcntIntCollab\",\"acadStaffStdnt\",\"acadStaffInt\",\"avrgPrcnt\",\"norm\",\"ncicountry\",\"avrgCitations\",\"location\",\"doctoral\",\"doctoralUndergrad\",\"docsCited\",\"esi\",\"hindex\",\"highlyCitedPapers\",\"impactRelToWorld\",\"instIncome\",\"intCollaborations\",\"jNCI\",\"level\",\"type\",\"papers\",\"papersInt\",\"resIncome\",\"resIncomeInd\",\"resReputGlob\",\"stateProvice\",\"stdntInt\",\"teachingReput\",\"the\",\"timesCited\",\"wosDocuments\",\"key\",\"seqNumber\",\"hasProfile\"]}"
headers = {
    'accept': "application/json, text/plain, */*",
    'origin': "https://incites.thomsonreuters.com",
    'accept-language': "en",
    'user-agent': "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.181 Safari/537.36",
    'content-type': "application/json;charset=UTF-8",
    'referer': "https://incites.thomsonreuters.com/",
    'accept-encoding': "gzip, deflate, br",
    'cookie': "_ga=GA1.2.1988740908.1527559100; _gid=GA1.2.501115603.1527559100; USERNAME=\"hazelnut@sjtu.edu.cn\"; CUSTOMER_NAME=\"SHANGHAI JIAO TONG UNIV\"; E_GROUP_NAME=\"IC2 Platform\"; SUBSCRIPTION_GROUP_ID=\"242426\"; SUBSCRIPTION_GROUP_NAME=\"SHANGHAI JIAO TONG UNIVERSITY_Benchmarking\"; CUSTOMER_GROUP_ID=\"93841\"; ROAMING_DISABLED=\"false\"; ACCESS_METHOD=\"UNP\"; _gat=1; JSESSIONID=CCCA0D0F2C717AE594199E625E609A41; PSSID=\"H1-D90zu5x2BmPnPHjHjHvjaqVeYZmjO3aaIm-18x2d4Z4Z6kWVnx2B4PrQKQ8eH2vwx3Dx3DyyZOoXUx2Fgq0uvKP57JBruQx3Dx3D-iyiHxxh55B2RtQWBj2LEuawx3Dx3D-1iOubBm4x2FSwJjjKtx2F7lAaQx3Dx3D\"; SECCONTEXT=a4e127693e537ca470fa3e0ca0932398b8ff9449dbef958eef72dc7b3e216f14188ba490ba80ccef20a295c4b4f072f8079c62c19d2bd961d0fc7ae7dbbb717d0bf5a12b86d8bcc7967ac9ae9b2151dea64ff8dec8b2bad7b005178896bb1fa2d28fb106d0f8b5e80b08c5e89e329a22697db49e7326dae37a3599b6dbf5292b15386d0daf34dc879dc30a4d5cca698e6d17ca9040456bc316a8cb1f0540fdf65b1c56f90e37f07ceff37948e71d89e09ffc153fa152d01172b362de9926463cc73c88d6d230bb3a9100b406fa87115772ed5cc901422a45148e61dc78312ca062da4f1f4cdbd11bfb5f2d6b3baaacf13fb6b99347b03e389afe5edf02c64e2389f07bed70f14227fe6d2c7b725b1e7326e0d64ea5c11089a1f47ca1624de53308f88c25310c3aeb932a4222a6b0690dee6176c77e2e031d82a960f74aa3184e32ad877d18f857450e46121206c24fbd5b01e8f75597f643694825c760ff5532711abf2925fdcfa41416484408bc4f70cce1bf5a050dabcad23c9661edbc16878ce7aa3a97cad0706392d4958a5560aa603d845657c3a19f83ff187d238d5ee9b65856884bde4c43d7b35419e6c55bd1e05c4ca9e2f84819ed72c80888978543f584088be6232c7c87671138020da1ccf9755008df263d7b3e573f61f3e0525c64c27812bd3621ab6a28da05466c45cce86d2cfeb7d456558335b82702af7111849529565e61d71ae16237402bcb8093; JSESSIONID=CB9C33241E20B71A0D3E9C7258989EF3; _ga=GA1.2.1988740908.1527559100; CUSTOMER_NAME=\"SHANGHAI JIAO TONG UNIV\"; ROAMING_DISABLED=\"false\"; _gid=GA1.2.242436379.1527747230; USERNAME=\"hazelnut@sjtu.edu.cn\"; PSSID=\"H2-2NqOi4YC6eL5cOB9VqufdMazkwod2Qnx2B-18x2dUE4r2u4PqmoPrQKQ8eH2vwx3Dx3DKh3mcv4TPYtO7vsgn6Jz8wx3Dx3D-iyiHxxh55B2RtQWBj2LEuawx3Dx3D-1iOubBm4x2FSwJjjKtx2F7lAaQx3Dx3D\"; E_GROUP_NAME=\"IC2 Platform\"; SUBSCRIPTION_GROUP_ID=\"242426\"; SUBSCRIPTION_GROUP_NAME=\"SHANGHAI JIAO TONG UNIVERSITY_Benchmarking\"; CUSTOMER_GROUP_ID=\"93841\"; ACCESS_METHOD=\"IP_ROAMING\"; SECCONTEXT=cc671d6c5028f4334aa537f7939cb6cda8548a0ba88db1d853522d889104a8bdd037b66fadb65699cdc09ad5721df7709148583b75fbde730a838ab59b7e9fcbe91a90c9eafadae11896f4318caafd54fbcf685bc02172879dd6dbec3273ed5c6ed8b6891f75e32768ac25ccbd6a13e24b76888e276c14c8b1576657543b0b5d935c3e1df429032ab983e071d69dcd930a1e5f8d37d0942d683467bd60382e255a486fdfaa26aae87146efbbb9a91700a7a0cd3418c9f7aed3ad1fc4df6a3f06956a1ce3c38d780d2deccd4c06a1a004e85675e73fa43ff542eedd77a6937d0c21a6d9019774d95a10f4a9f96fd5686e8135f0c780df6cd0e03c3dc6a0943864a2046ff87b6b18fcfbf44e2ffdf3729c0947113ccb94f98afe1f998978d0875b803792bd23e0cd0a69a05d107f247d41163f125574e3c84531808be16429e9c7e9dadc12f3b25f68f97234d9a6275c3f33fa82d56fa8b1ff006e4f55966c2320ddec24a50fee528cc25e17165becd27e28528a089ae79804e737f50579972afbdd5f5fd033ae1136193504d1b4bfc7f2dc746a83cf9d7ea2ccaa720bcf6d279a15fb01153d4ac579268ee3ab8d5016749a929e553b7ea6a5c1accc91007bb68c3d466e0137816757503543e5dd931a9dcd2afb44f2f22906e106bd8db8232b838698102c067fad75e90dc35e390827015d8cdb7363341d4db9b1f57e20e5fd8d722d17628376a07e22b5e9fb1c091a62; _gat=1",
    'cache-control': "no-cache",
    'postman-token': "6261ce99-bab9-c137-daf8-d90ac97da178"
    }

response = requests.request("POST", url, data=payload, headers=headers)
content = response.text
#print content
itemList = content.split('},{')
TOTAL = 1
while itemList[TOTAL][1:7] == "doctor":
    TOTAL += 1
print ("totally " + str(TOTAL) + ' institutions.')

print len(os.listdir(curPath))
print TOTAL

#while (len(os.listdir(curPath)) < TOTAL):
while (len(os.listdir(curPath)) < 1000):
    varLock = threading.Lock()
    urllock = threading.Lock()
    q = Queue.Queue()
    for index in range(TOTAL):
        q.put(index)
    taskList = []
    for i in range(NUM):
        t = threading.Thread(target=work, args=(i,))
        t.setDaemon(True)
        try:
            t.start()
            taskList.append(t)
        except Exception, e:
            print str(i) + 'the institution thread failed'
            q.put(i)
    for task in taskList:
        task.join()

if option ==1:
    myfile = open("Paper Collection.csv", 'wb')
else:
    myfile = open("Research Area Info Collection.csv", 'wb')
csv_writer = csv.writer(myfile, dialect="excel")
if option ==1:
    TitleLine = ["PaperID", "Location", "Institution","Accession Number","DOI","Pubmed ID","Article Title",
                 "Link","Authors","Source","Research Area","Volume","Issue","Pages","Publication Date",
                 "Times Cited","Journal Expected Citations","Category Expected Citations",
                 "Journal Normalized Citation Impact","Category Normalized Citation Impact",
                 "Percentile in Subject Area","Journal Impact Factor"]
else:
    TitleLine = ["ID","Institution","Research Area", "Rank", "Web of Science Documents", "Category Normalized Citation Impact",
                 "Times Cited",
                 "% Docs Cited", "% Documents in Q1 Journals", "% Documents in Q2 Journals",
                 "% Documents in Q3 Journals", "% Documents in Q4 Journals", "% Documents in Top 1%",
                 "% Documents in Top 10%", "% Highly Cited Papers", "% Hot Papers", "% Industry Collaborations",
                 "% International Collaborations", "Average Percentile", "Citation Impact",
                 "Documents in JIF Journals", "Documents in Q1 Journals", "Documents in Q2 Journals",
                 "Documents in Q3 Journals", "Documents in Q4 Journals", "Highly Cited Papers",
                 "Impact Relative to World", "International Collaborations", "Journal Normalized Citation Impact"]
print "Writing csv version."
csv_writer.writerow(TitleLine)
paperFiles = os.listdir(curPath)
IDCounter = 1
for FILE in paperFiles:
    mark = 1
    FILE = open(curPath + "\\" + FILE)
    csv_reader = csv.reader(FILE)
    for row in csv_reader:
        if mark == 1:
            mark = 0
            continue
        if option ==1:
            row[0] = IDCounter
            if row[7].find('gateway') == -1:
                row[6] = row[6] + '(' + row[7] + ')'
                row = row[0:7] + row[8:]
            if row[5][-1] == "\"":
                row[4] = row[4] + row[5]
                row[4] = row[4][1:-1]
                comma = row[6].find(",")
                row[5] = row[6][:comma]
                row[6] = row[6][comma + 1:]
            if row[6][0] == "\"":
                row[6] = row[6][1:]
            if len(row) == 21:
                row = row[:8] + ['n/a'] + row[8:]
            for index in range(len(row)):
                if row[index] == 'wosUrl\'':
                    print 'found one!'
                    row = row[:index] + row[index + 1, :]
        else:
            row = [IDCounter] + row
        csv_writer.writerow(row)
        IDCounter += 1
    FILE.close()
myfile.close()
print "csv version done. Writing access version."
if option == 1:
    write_db_for_PC(IDCounter/900000 + 1)
    print "access version done."
else:
    write_db_for_RA()
    print "access version done."
if os.path.exists("Paper Collection.csv"):
    os.remove("Paper Collection.csv")
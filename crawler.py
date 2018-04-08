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
            item = '\"' + item
            item = item + '\"'
            sql = sql + item
            sql = sql + ','
        sql = sql[0:len(sql) - 2]
        sql = sql + "\")"
        try:
            conn.execute(sql)
        except Exception as e:
            #print(e)
            print line
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
        'user-agent': "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.186 Safari/537.36",
        'accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
        'referer': "https://incites.thomsonreuters.com/",
        'accept-encoding': "gzip, deflate, br",
        'accept-language': "en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7",
        'cookie': "JSESSIONID=9081AE8656BD8511AA1459DBD9A2C5AA; _ga=GA1.2.289101491.1519811490; _gid=GA1.2.1819402626.1519811490; USERNAME=\"hazelnut@sjtu.edu.cn\"; PSSID=\"A2-sJAf7CXZdbxxSfhBe4EcL3j4Z6qZTslnxx-18x2dwx2Fx2FOBEgWmdTsoix2BefrHXeAx3Dx3DgI5YeJXOKqgBLntQx2BeOjJQx3Dx3D-9vvmzcndpRgQCGPd1c2qPQx3Dx3D-wx2BJQh9GKVmtdJw3700KssQx3Dx3D\"; CUSTOMER_NAME=\"SHANGHAI JIAO TONG UNIV\"; E_GROUP_NAME=\"IC2 Platform\"; SUBSCRIPTION_GROUP_ID=\"242426\"; SUBSCRIPTION_GROUP_NAME=\"SHANGHAI JIAO TONG UNIVERSITY_Benchmarking\"; CUSTOMER_GROUP_ID=\"93841\"; ROAMING_DISABLED=\"false\"; ACCESS_METHOD=\"UNP\"; _gat=1; SECCONTEXT=fc71f218d7503f6128861dd9f322ed5e518b3884ab470d2d10e5598d1339457780bd2b417b8e74387b9f7bc72ae1a5a1350d835ed22c582e0a254c1af22bb59bba796a57482a24f77805ff3faba033a830e6d911835cb0b2ebc2903449a64d55b7775785ce2606fd857d510a624d633be02125cfe0c88e6174f5e4a5d1662e8ad3d07a8daaf3f54383bfc2be2a15878b0e5c1062b7c54f9d46164a6809ab80dd537b1891e24e9cb5bcae2c0551b63a5a94eec988a5ec5b0e48bf59198a424c2d8bafa4e44c6a33e39af18c8710784d4438b5ec4ff9409b056f6540b0d15d25f94bdfefc7e965073955cd3599af77403e42241eca41354b3f9ce67b30176f2356d1812dfe4fd070b0420f20bc33f1867348e646938e4db6b3309ebfbdd7d30eb82c50e56e2e82badc21ba6f2e5c5c6b8021faf86c2b446e22df2694f57b27fadef213dd832e0e169360af498e875456120a810fd1030dfecea440887517d049b9df85a4d80fc426a47f85247808acea02e5e5cc439b0e1dd235b379f245c2460d6972dd225ad54787f505cffc5ec7736923db4f98aad722dff20fddc1f2de16a6",
        'cache-control': "no-cache",
        'postman-token': "87a92258-ab38-50b3-1184-e9230093861e"
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

def getMoreinfo(key, num, skip, name, number):
    url = "https://incites.thomsonreuters.com/incites-app/drilldowns/0/organization/dbd_39/data/export/csv"
    #querystring = {"skip": skip, "sortBy": "cites", "sortOrder": "desc", "take": num,"fileName": name, "key":key}
    querystring = {"skip":skip, "sortBy": "cites", "sortOrder": "desc", "take": "10000", "fileName": name, "key": key}
    # payload = "{\"take\":10000000,\"skip\":0,\"sortBy\":\"timesCited\",\"sortOrder\":\"desc\",\"filters\":{\"location\":{\"is\":[\"CHINA MAINLAND\",\"TAIWAN\",\"MACAU\",\"HONG KONG\"]},\"personIdTypeGroup\":{\"is\":\"name\"},\"personIdType\":{\"is\":\"fullName\"},\"schema\":{\"is\":\"Essential Science Indicators\"},\"articletype\":{\"is\":[\"Article\",\"Letter\",\"Review\"]},\"period\":{\"is\":[2007,2017]}},\"pinned\":[],\"indicators\":[\"orgName\",\"rank\",\"percentCited\",\"prcntDocsIn99\",\"prcntDocsIn90\",\"prcntHighlyCitedPapers\",\"prcntHotPapers\",\"prcntIndCollab\",\"prcntIntCollab\",\"acadStaffStdnt\",\"acadStaffInt\",\"avrgPrcnt\",\"norm\",\"ncicountry\",\"avrgCitations\",\"location\",\"doctoral\",\"doctoralUndergrad\",\"docsCited\",\"esi\",\"hindex\",\"highlyCitedPapers\",\"impactRelToWorld\",\"instIncome\",\"intCollaborations\",\"jNCI\",\"level\",\"type\",\"papers\",\"papersInt\",\"resIncome\",\"resIncomeInd\",\"resReputGlob\",\"stateProvice\",\"stdntInt\",\"teachingReput\",\"the\",\"timesCited\",\"wosDocuments\",\"key\",\"seqNumber\",\"hasProfile\"]}"
    #curPayload1 = "params={\"filters\":{\"location\":{\"is\":[\"CHINA MAINLAND\",\"TAIWAN\",\"HONG KONG\",\"MACAU\"]},\"personIdTypeGroup\":{\"is\":\"name\"},\"personIdType\":{\"is\":\"fullName\"},\"schema\":{\"is\":\"Essential Science Indicators\"},\"articletype\":{\"is\":[\"Article\",\"Review\",\"Letter\"]},\"period\":{\"is\":[2007,2017]}},\"pinned\":[],\"dateInfo\":{\"exportDate\":\"2017-08-06\",\"wosDate\":\"2017-05-31\",\"deployDate\":\"2017-07-22\"}}"
    flag = payload.find("\"filters\"")
    flag1 = payload.find("pinned")
    curPayload = "params={" + payload[flag: flag1 + 11] + "\"dateInfo\":{\"exportDate\":\"2017-08-06\",\"wosDate\":\"2017-05-31\",\"deployDate\":\"2017-07-22\"}}"
    headers = {
        'origin': "https://incites.thomsonreuters.com",
        'upgrade-insecure-requests': "1",
        'content-type': "application/x-www-form-urlencoded",
        'user-agent': "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.186 Safari/537.36",
        'accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
        'referer': "https://incites.thomsonreuters.com/",
        'accept-encoding': "gzip, deflate, br",
        'accept-language': "en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7",
        'cookie': "JSESSIONID=9081AE8656BD8511AA1459DBD9A2C5AA; _ga=GA1.2.289101491.1519811490; _gid=GA1.2.1819402626.1519811490; USERNAME=\"hazelnut@sjtu.edu.cn\"; PSSID=\"A2-sJAf7CXZdbxxSfhBe4EcL3j4Z6qZTslnxx-18x2dwx2Fx2FOBEgWmdTsoix2BefrHXeAx3Dx3DgI5YeJXOKqgBLntQx2BeOjJQx3Dx3D-9vvmzcndpRgQCGPd1c2qPQx3Dx3D-wx2BJQh9GKVmtdJw3700KssQx3Dx3D\"; CUSTOMER_NAME=\"SHANGHAI JIAO TONG UNIV\"; E_GROUP_NAME=\"IC2 Platform\"; SUBSCRIPTION_GROUP_ID=\"242426\"; SUBSCRIPTION_GROUP_NAME=\"SHANGHAI JIAO TONG UNIVERSITY_Benchmarking\"; CUSTOMER_GROUP_ID=\"93841\"; ROAMING_DISABLED=\"false\"; ACCESS_METHOD=\"UNP\"; _gat=1; SECCONTEXT=fc71f218d7503f6128861dd9f322ed5e518b3884ab470d2d10e5598d1339457780bd2b417b8e74387b9f7bc72ae1a5a1350d835ed22c582e0a254c1af22bb59bba796a57482a24f77805ff3faba033a830e6d911835cb0b2ebc2903449a64d55b7775785ce2606fd857d510a624d633be02125cfe0c88e6174f5e4a5d1662e8ad3d07a8daaf3f54383bfc2be2a15878b0e5c1062b7c54f9d46164a6809ab80dd537b1891e24e9cb5bcae2c0551b63a5a94eec988a5ec5b0e48bf59198a424c2d8bafa4e44c6a33e39af18c8710784d4438b5ec4ff9409b056f6540b0d15d25f94bdfefc7e965073955cd3599af77403e42241eca41354b3f9ce67b30176f2356d1812dfe4fd070b0420f20bc33f1867348e646938e4db6b3309ebfbdd7d30eb82c50e56e2e82badc21ba6f2e5c5c6b8021faf86c2b446e22df2694f57b27fadef213dd832e0e169360af498e875456120a810fd1030dfecea440887517d049b9df85a4d80fc426a47f85247808acea02e5e5cc439b0e1dd235b379f245c2460d6972dd225ad54787f505cffc5ec7736923db4f98aad722dff20fddc1f2de16a6",
        'cache-control': "no-cache",
        'postman-token': "87a92258-ab38-50b3-1184-e9230093861e"
    }
    if threading.active_count() > 5:
        while number%49 != int(100*time.time())%49:
            sleep(float(float(3)/50*threading.active_count()))
    response = requests.request("POST", url, data=curPayload, headers=headers, params=querystring)
    content = response.text
    paperList = content.split('\n')
    paperList = paperList[1:]
    while len(paperList) > 0 and paperList[-1].find("WOS") == -1:
        paperList = paperList[:-1]
    if len(paperList) == 1 and paperList[0].find("WOS") == -1:
        print content
        return []
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
    paperList = []
    times = int(num)/10000
    for tm in range(times + 1):
        skip = tm*10000
        skip = str(skip)
        paperList+=getMoreinfo(key, num, skip, name, number)
    if len(paperList) < int(num) - 10:
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
                getInfo(key, num, name, location, number)
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

payload = "{\"take\":10000000,\"skip\":0,\"sortBy\":\"timesCited\",\"sortOrder\":\"desc\",\"filters\":{"

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
payload += "]}},\"pinned\":[],\"indicators\":[\"orgName\",\"rank\",\"percentCited\",\"prcntDocsIn99\",\"prcntDocsIn90\",\"prcntHighlyCitedPapers\",\"prcntHotPapers\",\"prcntIndCollab\",\"prcntIntCollab\",\"acadStaffStdnt\",\"acadStaffInt\",\"avrgPrcnt\",\"norm\",\"ncicountry\",\"avrgCitations\",\"location\",\"doctoral\",\"doctoralUndergrad\",\"docsCited\",\"esi\",\"hindex\",\"highlyCitedPapers\",\"impactRelToWorld\",\"instIncome\",\"intCollaborations\",\"jNCI\",\"level\",\"type\",\"papers\",\"papersInt\",\"resIncome\",\"resIncomeInd\",\"resReputGlob\",\"stateProvice\",\"stdntInt\",\"teachingReput\",\"timesCited\",\"wosDocuments\",\"key\",\"seqNumber\",\"hasProfile\"]}"
# payload = "{\"take\":10000000,\"skip\":0,\"sortBy\":\"timesCited\",\"sortOrder\":\"desc\",\"filters\":{\"location\":{\"is\":[\"CHINA MAINLAND\",\"TAIWAN\",\"MACAU\",\"HONG KONG\"]},\"personIdTypeGroup\":{\"is\":\"name\"},\"personIdType\":{\"is\":\"fullName\"},\"schema\":{\"is\":\"Essential Science Indicators\"},\"articletype\":{\"is\":[\"Article\",\"Letter\",\"Review\"]},\"period\":{\"is\":[2007,2017]}},\"pinned\":[],\"indicators\":[\"orgName\",\"rank\",\"percentCited\",\"prcntDocsIn99\",\"prcntDocsIn90\",\"prcntHighlyCitedPapers\",\"prcntHotPapers\",\"prcntIndCollab\",\"prcntIntCollab\",\"acadStaffStdnt\",\"acadStaffInt\",\"avrgPrcnt\",\"norm\",\"ncicountry\",\"avrgCitations\",\"location\",\"doctoral\",\"doctoralUndergrad\",\"docsCited\",\"esi\",\"hindex\",\"highlyCitedPapers\",\"impactRelToWorld\",\"instIncome\",\"intCollaborations\",\"jNCI\",\"level\",\"type\",\"papers\",\"papersInt\",\"resIncome\",\"resIncomeInd\",\"resReputGlob\",\"stateProvice\",\"stdntInt\",\"teachingReput\",\"the\",\"timesCited\",\"wosDocuments\",\"key\",\"seqNumber\",\"hasProfile\"]}"
headers = {
    'accept': "application/json, text/plain, */*",
    'origin': "https://incites.thomsonreuters.com",
    'accept-language': "en",
    'user-agent': "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.186 Safari/537.36",
    'content-type': "application/json;charset=UTF-8",
    'referer': "https://incites.thomsonreuters.com/",
    'accept-encoding': "gzip, deflate, br",
    'cookie': "JSESSIONID=D45FD0C55DEDDC7FAE58ABE9CC516F2B; _ga=GA1.2.953928054.1520065904; _gid=GA1.2.1302978877.1520065904; _gat=1; USERNAME=\"hazelnut@sjtu.edu.cn\"; PSSID=\"B1-Axx1O7bImcaVTmx2BdCikKjZyUmK8TRhMiX-18x2dzIUOwLWjURMPrQKQ8eH2vwx3Dx3DR0ZaHPdVeOJIGwx2F3RUd7zQx3Dx3D-YwBaX6hN5JZpnPCj2lZNMAx3Dx3D-jywguyb6iMRLFJm7wHskHQx3Dx3D\"; CUSTOMER_NAME=\"SHANGHAI JIAO TONG UNIV\"; E_GROUP_NAME=\"IC2 Platform\"; SUBSCRIPTION_GROUP_ID=\"242426\"; SUBSCRIPTION_GROUP_NAME=\"SHANGHAI JIAO TONG UNIVERSITY_Benchmarking\"; CUSTOMER_GROUP_ID=\"93841\"; ROAMING_DISABLED=\"false\"; ACCESS_METHOD=\"IP_ROAMING\"; SECCONTEXT=a76a6048f32c73978f1fb465cb57d4cab1a44d546aa2c1c3d5a1a4197406bf3e7bb1f73bfe9fabecbeb20c58b2f8825e4553f39ecf63cdddd09e99494d17e9a00fa6838dd96acb971682d1b79cd161ff539b671f2628e720ff81877a7178464b54d10700f23889d7d10824d113d2a39c5cb6ae2a583abbab43aa606638101ec29db0821838132c9028dc35352a4ba8bbecb7123af9db5c47a0222f8be8c8438977eb209cc133adf1e76580fe9820a96e0c8ea459611fcdff0c03ade56991a5aa12a8d8ecb495221af167ed2fb378176b19a01535cefa161de645f5865eda75c243cffc09fcb78be2e85fd3580979b3fe37704a6b2c061ae0c27e0ea6e96a07dfe7aaa1d467a1d096af0fb7cb11c0ea284ca02c81306a74b6e42040928567b448939a0df16834ad2e1f642f1c29fa8e232ee50d8b4a2895204c754ae3e0037c75b7a6469071dcb4187108e865403ee2d97513e26797ae4b35b10928f87fa25eb6d468df0770debc42d2ac6cc86a5fc010ee0a57a1df145081309ce209a3afae359a7bc05cb242683534d2773d06e1c6253ee4b118ac3cfbae9067d809df12cb54",
    'cache-control': "no-cache",
    'postman-token': "52d219d6-2dd5-6971-999f-724afa76d605"
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

while (len(os.listdir(curPath)) < TOTAL):
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
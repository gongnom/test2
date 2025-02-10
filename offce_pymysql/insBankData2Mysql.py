#-*- coding: utf-8 -*-
# 법인은행 데이터를 MySQL에 insert하기
# 잘 동작하고 있음. 
import sys, os
# sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))   #부모경로를 참조 path에 추가
# from Config import config             # Config/config.py 임포트
import pymysql
from openpyxl import load_workbook, Workbook
import datetime
import xlrd
import logging
import mydb
from myconfig import DIR, TEST_DB, Account_no, CENTER_DB, CENTER_DB_LOCAL
import re

def load_excel_file(file_name):
    # file_name[은행 데이터(xls file)]을 불러 계좌번호, 내용(리스트)을 반환함. 
    xlsfile  = os.path.join(DIR['bank'], file_name)
    wbs = xlrd.open_workbook(xlsfile)
    wss = wbs.sheet_by_index(0)
    nrows = wss.nrows # 행(가로줄 수)
    ncols = wss.ncols # 열(세로줄 수)
    data = [[wss.cell_value(i,j) for j in range(ncols)] for i in range(nrows) ]
    account_num = data[2][1]
    # 신한 7번째 부터 실제 내용
    del data[:7]

    return account_num, data

def loadxls2mysqldata(last_daytime):
    # 법인통장 xls파일 중 last_daytime 이후 데이터만 mysql 데이터로 반환 

    # 1. 불러올 법인 통장 파일이름 만들기 
    myday = datetime.date.today()
    file_name = myday.strftime("%y%m")+'bub.xls' #bub.xls로 바꿔도 됨 
    # file_name = '2103bub.xls' #bub.xls로 바꿔도 됨 
    
    # 2. 법인 통장 xls 파일에서 계좌번호, 데이터를 불러오기 
    acnum, mydata = load_excel_file(file_name)
    
    # 3. 불러온 법인통장 데이터 중 last_dayime 이후 데이터만 sql data로 변환하기
    mysqldata = []
    for i, v in enumerate(mydata):
        # bubin, gong mydata : 0일자, 1시간, 2적요(타행 ), 3출, 4입, 5내용, 6잔액, 7거래점, 8cms 코드
        md = mydata[i][0]
        mt = mydata[i][1] 
        dt = datetime.datetime(int(md[:4]), int(md[5:7]), int(md[8:]), int(mt[:2]), int(mt[3:5]), int(mt[6:]))
        if dt > last_daytime: 
            torf = 'F'
            if int(mydata[i][4]) == 0 : torf = 'T'
            mysqldata.append([
                acnum,                #0 account_num varchar(20) not null,
                mydata[i][0]+' '+mydata[i][1],         #1 daytime datetime not null, 
                mydata[i][5],         #2 contents varchar(30) not null, 
                int(mydata[i][3]),    #3 withdraw int default 0, 
                int(mydata[i][4]),    #4 deposit int default 0, 
                int(mydata[i][6]),    #5 balance int not null, 
                '',                   #6 bigo varchar(100), 
                torf                 #7 depos_tax char(1)
                ])
            print("insert bank : {} > {} mysql에 추가될 예정입니다.".format(dt, last_daytime))
        else:
            print("insert bank : {} <= {} 추가되지 않습니다.".format(dt, last_daytime))
    return mysqldata

def ins_data(mydb1):
    # mydb1=mydb.MyDB(cfg)
    # 1. mysql bank table에서 제일 마지막 데이터 중에 날짜시간을 불러오기 
    sql_last = "SELECT daytime FROM bank WHERE account_num = '100-029-410640' ORDER BY daytime DESC LIMIT 1"
    try:
        daytime = mydb1.run_query(sql_last)
        print("계좌번호 : {}, 마지막 데이터 거래일시 : {}".format(Account_no['bub'] , daytime[0][0]))
    except:
        print("mydb.ins_data : database load error")


    
    # 2. 법인통장 xls파일 중 last_daytime 이후 데이터만 mysql 데이터 가지고 오기
    try:
        mysql_data = loadxls2mysqldata(daytime[0][0])
    except:
        print("ins_data : error")
    # 3. 가지고 온 sql data를 bank table에 삽입 
    try:
        sql = "insert ignore into bank (account_num, daytime, contents, withdraw, deposit, balance, bigo, depos_tax) VALUES(%s, %s, %s, %s, %s,%s, %s, %s);"
        for i in mysql_data:
            # print(i)
            print(mydb1.run_query(sql, i))
            # make_case_day(mydb1, i)
    except:
        print("ins_data : datavase write error")
#def make_case_day(mydb1, i):
def make_case_day():
    mydb1=mydb.MyDB(TEST_DB)
    sql = "SELECT day, type, company, contents, party, fee, successway from mycase where party LIKE %s"
    i=('%곽도림%')
    a = mydb1.run_query(sql, i)
    print(a)

def make_tax_info(mydb1):
    # 법인계좌에서 F 되어 있는 내용을 불러와 guess한 다음 mytax table에 삽입하는 함수
    # mydb1=mydb.MyDB(cfg)

    # 법인계좌에서 F 되어 있는 내용을 불러오는 sql구문
    sql = "SELECT _id, daytime, contents, deposit FROM bank WHERE account_num = '100-029-410640' and depos_tax = 'f' ORDER BY daytime" # DESC"
    # 은행데이터    0      1         2        3
    data = mydb1.run_query(sql)
    
    if not data: 
        print("Make Tax info : mytax에 적을 내용이 없습니다.") # 법인계좌에 F가 없음.
        return 0,0
    else: 
        print("Make Tax info : 법인계좌 목록 중 'F' 목록을 mytax에 적습니다. ")

    datanum = len(data)
    fail_num = 0
    
    ######## guess start #############
    for i in data:
        # print(i)
        day = datetime.date(i[1].year, i[1].month, i[1].day)
        print("Make Tax info : 은행 데이터 {}".format(i))
        
        # 1. 사건 table에서 먼저 검색
        sql2 = "SELECT * FROM mycase WHERE (fee = %s and feeday = %s) or ( success = %s and successday = %s) ORDER BY day"
        print(i[3], day)
        mylist = mydb1.run_query(sql2, (i[3], day, i[3],  day))
        if mylist:
            print("Make Tax info : 사건 : {}".format(mylist))
            if len(mylist) > 1:
                for no, j in enumerate(mylist):
                    if j[3]=='중노위':
                        mytax_contents='구제 재심신청 사건'
                    elif j[3]=='지노위':
                        mytax_contents='구제 신청 사건'
                    elif j[3]=='노동청':
                        mytax_contents='진정 사건'
                    else:
                        mytax_contents=j[3]
                    print("{}. {} 입금".format(no, j[6]))
                    # mylist_no = input("Make Tax info : 여러개 있습니다.위에서 입금자를 찾아 번호를 적어주세요")
                    if i[2].strip() == j[6].strip():
                        print("{}번 인걸로 추측됨.".format(no))
                        if j[3] in ['의견서', '분석']:
                            contents = j[4] + ' ' + j[5] + ' ' + mytax_contents +' (' + j[6] + ') 비용'
                            compday = j[16]
                            bigo = j[3]
                        elif day == j[9]:
                            contents = j[4] + ' ' + j[5] + ' ' + mytax_contents + j[6] + ') 착수금'
                            compday = j[16]
                            bigo = '착수금'
                        else:
                            contents = j[4] + ' ' + j[5] + ' ' + mytax_contents +'(' + j[6] + ') 성공보수'
                            compday = j[17]
                            bigo = '성공보수'
                        tax_data = (i[0], j[0], day, i[3], j[4], j[6],  j[14], j[15], contents, compday, bigo)
                if not tax_data:
                    tax_data = (i[0], '', day, i[3], i[2], '',  'F', '', '알수없음', '', '?')
            else:
                if mylist[0][3]=='중노위':
                    mytax_contents='구제 재심신청 사건'
                elif mylist[0][3]=='지노위':
                    mytax_contents='구제 신청 사건'
                elif mylist[0][3]=='노동청':
                    mytax_contents='진정 사건'
                else:
                    mytax_contents=mylist[0][3]

                if mylist[0][3] in ['의견서', '분석']:
                    contents = mylist[0][4] + ' ' +  mylist[0][5] + ' ' + ' ' + mytax_contents + '(' + mylist[0][6] + ')' + ' 비용'
                    compday = mylist[0][16]
                    bigo = mylist[0][3]
                elif day == mylist[0][9]:
                    contents = mylist[0][4] + ' ' +  mylist[0][5] + ' ' +  ' ' + mytax_contents + '(' + mylist[0][6] + ')' + ' 착수금'
                    compday = mylist[0][16]
                    bigo = '착수금'
                else:
                    contents = mylist[0][4] + ' ' +  mylist[0][5] + ' ' + ' ' + mytax_contents + '(' + mylist[0][6] + ')' + ' 성공보수'
                    compday = mylist[0][17]
                    bigo = '성공보수'
                tax_data = (i[0], mylist[0][0], day, i[3], mylist[0][4], mylist[0][6],  mylist[0][14], mylist[0][15], contents, compday, bigo)
        
        # 사건(mycase) 테이블에서 없는 경우 -> 노조 테이블에서 찾음.
        
        else:
            print("Make Tax info : 사건 목록에 없어 노동조합으로 찾습니다 : ")
            print('i[2] : ', i[2])
            bankcontent = i[2]
            bankcontent = re.sub(r'\d+월|\d+', '', bankcontent) 
            print('i[2] : ',bankcontent)

            sql3 = "SELECT * FROM mynojo WHERE inputname like %s and fee = %s"
            ml = mydb1.run_query(sql3, ("%"+bankcontent+"%", i[3]))
            if not ml :
                sql3 = "SELECT * FROM mynojo WHERE inputname like %s"
                ml = mydb1.run_query(sql3, "%"+bankcontent+"%")
                print("Make Tax Info : 노조이름으로만 찾았습니다. 금액이 다를 수 있으니 참고 하기 바람.")
                
            """if len(i[2]) == 7:
                j = list(i)
                # 노조 테이블에서 검색
                sql3 = "SELECT * FROM mynojo WHERE inputname like %s and fee = %s"
                print('#######', j[2])
                j[2] = "%"+j[2][2:]+"%"
                ml = mydb1.run_query(sql3, (j[2], i[3]))
            else:
                # sql3 = "SELECT * FROM mynojo WHERE inputname = %s and fee = %s"
                # ml = mydb1.run_query(sql3, (i[2], i[3]))
                sql3 = "SELECT * FROM mynojo WHERE inputname like %s"
                print(i[2])    
                ml = mydb1.run_query(sql3, i[2])
                print("Make Tax Info : 노조이름으로만 찾았습니다. 금액이 다를 수 있으니 참고 하기 바람.")
            """
            # print(ml)
            # tax_data = "잠시후"
            #          (id, case_num, inputday, input, name, leader,    cashor,        info,          contents, completeday, bigo)
            null = None
            if ml:
                nojo_sql = "UPDATE mynojo SET lastday = %s WHERE id = %s"
                nojo1 = mydb1.run_query(nojo_sql, (day, ml[0][0]))
                if nojo1 != 0:
                    nojo_sql2 = "SELECT month, fee from mynojo where id = %s"
                    nojo2, fee2 = mydb1.run_query(nojo_sql2, ml[0][0])[0]
                    if nojo2 == 12:
                        month1 = 1
                    elif int(fee2)*2 == int(i[3]):
                        month1 = 2
                    else:
                        month1 = nojo2+1
                    nojo_sql3 = "UPDATE mynojo SET month = %s WHERE id = %s"
                    nojocheck = mydb1.run_query(nojo_sql3, (month1, ml[0][0]))
                    if nojocheck:
                        print("Mak0e Tax info : Updated nojo table as below-- ")
                        print("Make Tax info : {} 노조 자문료 {}월분 {}원 입금.".format(ml[0][1], month1, i[3]))
                    else:
                        print("\n----------------\n노조테이블 마지막 입금된 달이 변경되지 않았습니다.\n--------------------\n")
                    month_str = "{} {}월 자문료".format(ml[0][1], str(month1))
                else:
                    nojo_sql2 = "SELECT month from mynojo where id = %s"
                    nojo3 = mydb1.run_query(nojo_sql2, ml[0][0])[0][0]
                    month_str = "{} {}월 자문료".format(ml[0][1], str(nojo3))
                    print("Make Tax info : 노조 자문료 : 입금된 것이 없음.")
                tax_data = (i[0], ml[0][0], day, i[3], ml[0][1], ml[0][2],  ml[0][6], ml[0][7], month_str, null, 'nojo')

            else:
                print("노동조합에도 일치하는 데이터가 없습니다. 구분에 'F'로 체크합니다")
                if i[2]=='박성우':
                    tax_data = (i[0], 0, day, i[3], i[2], '박성우',  'F', '1234', '대표 전입금', null, '전입금')
                else:    
                    tax_data = (i[0], null, day, i[3], i[2], null,  'F', '1234', '알수없음', null, '?')
                # tax_data = None
                fail_num +=1
        
        ######## guess end #############
            
        print("Make Tax info : Tax Data : {}".format(tax_data))

        ############### mytax에 세금처리 입력해주기 시작 #####################################
        sql4 = """insert ignore into mytax
            (id, case_num, inputday, input, name, leader, cashor, info, contents, cmplday, bigo)
            VALUES(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s);"""
        # mytax에 세금처리 sql을 실행 
        flag_mytax_insert = mydb1.run_query(sql4, tax_data)
        print("Make Tax info : {} is inserted into mytax.".format(flag_mytax_insert))
        
        # mytax에 세금처리가 잘 입력된 경우 bank 디비에 T를 입력해주고, 아닌 경우 그 행을 나타내줌.
        if flag_mytax_insert == 1:
            # mytax에 잘 입력 된 경우

            # bank database에 'T' 입력
            usql = "update bank set depos_tax = 'T' where _id = %s"
            flag_bank_update = mydb1.run_query(usql, tax_data[0])
                                               

            # bank database에 'T'가 잘 입력되었는지 확인하기 위해 select 함.
            sqldep = "SELECT depos_tax from bank where _id = %s"
            sign_depos_tax = mydb1.run_query(sqldep, tax_data[0])[0][0]
            if flag_bank_update == 1:
                print("Make Tax info : bank table update되었습니다. bank table depos_tax의 값은 {}입니다".format(sign_depos_tax))
            else :
                print("Make Tax info : bank table의 depas_tax가 update되지 않았습니다!!")
                print("Make Tax info : bank table depos_tax의 값은 {}입니다".format(sign_depos_tax))
        else :
            # mytax에 잘 입력 되지 않은 경우 어떤 것인지 나타내주기위해. 
            sqldep = "SELECT depos_tax from bank where _id = %s"
            sign_depos_tax = mydb1.run_query(sqldep, tax_data[0])[0][0]
            print("Make Tax info : bank table의 depas_tax가 update되지 않았습니다!!")
            print("Make Tax info : bank table depos_tax의 값은 {}입니다".format(sign_depos_tax))
        print("\n")
        ############### mytax에 세금처리 입력해주기 끝 #####################################
    
    
    return datanum, fail_num


def main():
    # test DB인지 실제 db인지 확인 
    db = mydb.MyDB(TEST_DB)
    # ins_data(db)
    total, fail = make_tax_info(db)
    if fail=='0':
        print("전체 {}개 중 insert {}개가 실패하였습니다.".format(total, fail))
    else:
        print("전체 {}개 모두 insert 성공하였습니다.".format(total, fail))



if __name__ == "__main__":
    main()
    # make_case_day()
    # print(loadxls2mysqldata())

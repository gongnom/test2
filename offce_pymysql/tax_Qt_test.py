
import mydb
from myconfig import DIR, Account_no, Center_DB_Interner 
from insBankData2Mysql import ins_data, make_tax_info
import sys, os
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtCore import *
import datetime
import invoice
import time
import calendar

form_class = uic.loadUiType(DIR['base']+"//pymysql_bank//tax_Qt.ui")[0]


class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        # db 이름 확인
        db =  Center_DB_Interner
        
        self.db = mydb.MyDB(db)

        # download 폴더의 xls file을 bankdata 폴더로 이동하고 년월bub.xls로 만들기
        self.down2bank_dir_move()
        
        # 은행 엑셀 자료 db에 올리기
        try:
            ins_data(self.db) # bank table에 삽입만 함.   
        except:
            print("처리할 법인은행 파일이 없거나 내용이 없습니다.")
        
        # 세금처리 디비에 올라가지 않은 내역
        self.setTableWidgetData()
        
        self.pushButton.clicked.connect(self.makeCashTax) # 미처리 영수 내역 보기
        try:
            
            self.pushButton_4.clicked.connect(self.issue_cashTax) # 현금영수증 버튼
        except:
            self.textBrowser.append("현금영수증 발급 중 실패하였습니다. 다시 시도해 주세요.")
            
        self.pushButton_3.clicked.connect(self.issue_billTax)# 세금계산서 버튼

        # QDialog 설정
        self.msg = QMessageBox()

    
    
    def down2bank_dir_move(self):
        import xlrd
        import shutil

       
        bankdata_dir = DIR['bank']
        dl_dir = DIR['download']

       
        myfile = None 
        
        file_time = 0
        for file in os.listdir(dl_dir):
            file = os.path.join(dl_dir, file)
            try:
                wbs = xlrd.open_workbook(file)
                wss = wbs.sheet_by_index(0)
                acc_no = wss.cell_value(2,1)
            except:
                acc_no = None
            if acc_no == Account_no['bub'] and file.endswith('.xls') and os.path.basename(file).startswith('신한은행'):
                
                if file_time < os.path.getmtime(file):
                    file_time = os.path.getmtime(file) 
                    myfile = file

       
        if myfile:
            self.textBrowser.append("다운로드 폴더에 {} 파일이 존재합니다. 법인 엑셀파일로 보입니다.".format(myfile))

         

            myday = datetime.date.today()
            file_name = myday.strftime("%y%m")+'bub.xls' #bub.xls로 바꿔도 됨 
            destination = os.path.join(bankdata_dir, file_name)
            try: 
                shutil.move(myfile,destination)
                self.textBrowser.append("상기 파일을 {} 이동하였습니다.".format(destination))
            except:
                self.textBrowser.append("상기 파일 이동을 실패하였습니다. 수동으로 해보아요~")
        else:
            self.textBrowser.append("다운로드 폴더에 법인 은행 파일이 없습니다. 다시 다운로드하여 주세요.")



   
    def issue_billTax(self):
        
        import xlwt
        import platform

        sql = "SELECT id, inputday, input, name, leader, info, bigo, case_num FROM mytax WHERE cmplday is null and cashor = 'i'"
        #              0,        1,    2,    3,    4,      5     6       7
        cashdata = self.db.run_query(sql)
        self.textBrowser.clear()
        print(cashdata)
        
        
        # print("세금계산서 처리를 위해 연월tax.xls를 만들었습니다.")
        workbookw = xlwt.Workbook(encoding='utf-8')  # utf-8 인코딩 방식의 workbook 생성
        worksheetw = workbookw.add_sheet('sheet1')  # 시트 생성

        # 2번째 행부터 데이터를 입력합니다.
        for num, l in enumerate(cashdata):
            taxdate = l[1].strftime("%Y%m%d")
            taxday = l[1].strftime("%d")
            gong = round(l[2]/1.1)
            buga = l[2]-gong
            if l[6] == 'nojo': 
                contents = "월 자문료"
                email_sql = "SELECT email FROM mynojo WHERE id = %s"
                ur_email = self.db.run_query(email_sql, l[7])[0][0]
                print(ur_email)
                # if not ur_email:
                    # ur_email=""

            else:
                contents = l[6]
                ur_email = ""
            worksheetw.write(6+num, 0, '01') # 세금계산서 종류 01 : 일반
            worksheetw.write(6+num, 1, taxdate) # 작성일자
            worksheetw.write(6+num, 2, '100000000') 
            worksheetw.write(6+num, 4, '법인')
            worksheetw.write(6+num, 5, 'dkdkdkdkdk')
            worksheetw.write(6+num, 9, 'ggggg@gmail.com')
            worksheetw.write(6+num, 10, l[5]) # 등록번호
            worksheetw.write(6+num, 12, l[3]) # 상호
            worksheetw.write(6+num, 13, l[4]) # 대표
            worksheetw.write(6+num, 17, ur_email) # e-mail
            worksheetw.write(6+num, 19, gong) # 공급가액
            worksheetw.write(6+num, 20, buga) # 세액
            worksheetw.write(6+num, 22, taxday) # 일자
            worksheetw.write(6+num, 23, contents)# 품목
            worksheetw.write(6+num, 27, gong) # 공급가액
            worksheetw.write(6+num, 28, buga) # 세액
            worksheetw.write(6+num, 58, '01') # 01 :영수, 02 : 청구
            
            print("{} : tax.xls 파일에 입력하였습니다.".format(l[3]))
            
            self.textBrowser.append("{} : tax.xls 파일에 입력예정입니다.".format(l[3]))
            #
            try:
                # update cmplday
                ut_sql = "update mytax set cmplday = %s where id = %s"
                update = self.db.run_query(ut_sql, (l[1], l[0]))
                if update:
                    print("{} : update mytax set comlday = {}".format(l[3], l[1]))
                    self.textBrowser.append("{} : update mytax set comlday = {}".format(l[3], l[1]))
                else:
                    print("{} : not update mytax".format(l[3]))
                    self.textBrowser.append("{} : not update mytax".format(l[3]))

            except:
                reply = QMessageBox.question(self, 'Message', "update mytax error : id : {}, 회사 : {}".format(l[0], l[3]), QMessageBox.Ok)
                # if reply == QMessageBox.Ok: self.msg.close()
                print("update mytax error : id : {}, 회사 : {}".format(l[0], l[3]))
            
        myday = datetime.datetime.today().strftime("%y%m%d")
        file_name=f"tax{myday}.xls"
        xlsfile  = os.path.join(DIR['bank'], file_name)
        workbookw.save(xlsfile)
        self.textBrowser.append('파일을 만들었습니다.')
        
      
        invo2 = invoice.IssueInvoice()
        self.textBrowser.clear()
        self.textBrowser.append(invo2.dt.strftime("%Y.%m.%d.%H:%M:%S") + " - 세금계산서 발행")
        
        self.textBrowser.append(invo2.login())

    def makeCashTax(self):
        # 미처리 영수 내역 보기
        # 법인계좌에서 F 되어 있는 내용을 불러와 guess한 다음 mytax table에 삽입
        make_tax_info(self.db) # insBankData2Mysql 모듈  
        
        sql = """SELECT DATE_FORMAT(inputday, '%Y-%m-%d'), format(input , 0),contents,  cashor 
          FROM mytax WHERE cmplday is null ORDER BY inputday DESC"""
        casedata = self.db.run_query(sql)
        self.tableWidget_2.setRowCount(len(casedata))

        for i, v in enumerate(casedata):
            for j, val in enumerate(v):
                # print(i, j, val)
                item = QTableWidgetItem(val)
                self.tableWidget_2.setItem(i, j, item)

        self.tableWidget_2.resizeColumnsToContents()
        # self.tableWidget_2.resizeRowsToContents()
        mytax_sum, bank_sum, my_ret_str = self.checkup()
        self.textBrowser.clear()
        self.textBrowser.append(self.cashor_tax_check())
        self.textBrowser.append("세금 디비 총액 : {}\n은행 총액 : {}\n{}".format(mytax_sum, bank_sum, my_ret_str))

    def checkup(self):
        # bank db와 mytax db 비교
        today = datetime.date.today()
        start = today.replace(day=1).strftime("%Y-%m-%d")
        end = "{}-{}-{}".format(today.year, today.month, calendar.monthrange(today.year, today.month)[1])
        sql1 = "SELECT format(sum(input) , 0) FROM mytax WHERE date(inputday) between %s and %s"
        am1 = self.db.run_query(sql1, (start, end))
        sql2 = "SELECT format(sum(deposit) , 0) FROM bank WHERE date(daytime) between %s and %s"
        am2= self.db.run_query(sql2, (start, end))
        if am1[0][0] == am2[0][0]: return am1[0][0],am2[0][0], "Check up : 합이 같습니다"
        else: am1[0][0],am2[0][0], "Check up : 합이 같지 않습니다ㅠㅠ"

    def cashor_tax_check(self):
        # mytax db의 현금영수증, 세금계산서, 필요없는 사항 총액 알려주기
        today = datetime.date.today()
        start = today.replace(day=1).strftime("%Y-%m-%d")
        end = "{}-{}-{}".format(today.year, today.month, calendar.monthrange(today.year, today.month)[1])
        sqli = "SELECT format(sum(input) , 0) FROM mytax WHERE cashor = 'i' and date(inputday) between %s and %s" 
        sqlc = "SELECT format(sum(input) , 0) FROM mytax WHERE cashor = 'c' and date(inputday) between %s and %s" 
        sqlf = "SELECT format(sum(input) , 0) FROM mytax WHERE cashor = 'f' and year(inputday) = %s and month(inputday) = %s" 
        ami= self.db.run_query(sqli, (start, end))
        amc= self.db.run_query(sqlc, (start, end))
        amf= self.db.run_query(sqlf, (str(today.year), str(today.month)))
        return """Check up : {}월 \n현금영수증 처리(예정 포함) 총액은 {}원입니다.\n세금계산서 처리(예정 포함) 총액은 {}원입니다. \n세금처리가 필요하지 않은 총액은 {}원 입니다.""".format(today.month, amc[0][0], ami[0][0], amf[0][0])
    
    
    def setTableWidgetData(self):
        # 세금처리 디비에 올라가지 않은 은행 입금 내역
        sql = """SELECT DATE_FORMAT(daytime, '%Y-%m-%d'), contents, format(deposit , 0), depos_tax 
          FROM bank WHERE depos_tax = 'F' ORDER BY daytime DESC"""
        try:
            casedata = self.db.run_query(sql)
            rowcount = len(casedata)
        except:
            casedata = []
            rowcount = 1
        self.tableWidget.setRowCount(rowcount)

        for i, v in enumerate(casedata):
            for j, val in enumerate(v):
                item = QTableWidgetItem(val)
                self.tableWidget.setItem(i, j, item)


    def issue_cashTax(self):
        # 현금영수증 처리(국세청) 및 현금영수증 처리된 날을 기록(mytax)
        sql = "SELECT id, case_num, input, name, info, bigo, contents FROM mytax WHERE cmplday is null and cashor = 'c'"
    #                 0,        1,    2,    3,   4,    5,      6     
        cashdata = self.db.run_query(sql)
        if not cashdata: 
            self.textBrowser.clear()
            self.textBrowser.append("현금영수증 처리할 게 아무것도 없음.")
            print("현금영수증 처리할 게 아무것도 없음.")
            reply = QMessageBox.question(self, 'Message', '현금영수증 처리할 게 아무것도 없음.', QMessageBox.Ok)
            if reply == QMessageBox.Ok: self.msg.close()
            return False
        invo = invoice.IssueInvoice()
        # print(invo.dt.strftime("%Y.%m.%d.%H:%M:%S") + " - issue Cash")
        self.textBrowser.clear()
        self.textBrowser.append(invo.dt.strftime("%Y.%m.%d.%H:%M:%S") + " - issue Cash")
        self.textBrowser.append(invo.login(1))
        # print(invo.login(1))
        
        for v in cashdata:
            print(v)
            # [cash or invoice, v[6] : contents, v[2] : input 금액, v[4] : info 전화번호] 
            if v[4].startswith('10'): v[4]='0'+v[4]
            # print(invo.cash_info(v[6], v[2], v[4])) 
            self.textBrowser.append(invo.cash_info(v[6], v[2], v[4]))
            try:
                myday = datetime.date.today()
                ut_sql = "update mytax set cmplday = %s where id = %s"
                # print(mydb1.run_query(ut_sql, (myday, v[0])))
                return_ut = self.db.run_query(ut_sql, (myday, v[0]))
                if return_ut == 1:
                    self.textBrowser.append("mytax db에 update가 잘 되었습니다.")
            except:
                self.textBrowser.append("update mytax error : id : {}, 회사 : {}".format(v[0], v[3]))
                # print("update mytax error : id : {}, 회사 : {}".format(v[0], v[3]))
            
            time.sleep(2)
        # print(invo.exit())
        self.textBrowser.append(invo.exit())
        

if __name__ == "__main__" :
    #QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv) 

    #WindowClass의 인스턴스 생성
    myWindow = WindowClass() 

    #프로그램 화면을 보여주는 코드
    myWindow.show()

    #프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()
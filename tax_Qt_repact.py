# Model: 데이터 및 비즈니스 로직
import os
import shutil
import datetime
import xlrd
from mydb import MyDB
from myconfig import DIR, Account_no, CENTER_DB_LOCAL
from insBankData2Mysql import ins_data, make_tax_info

class BankModel:
    def __init__(self):
        self.db = MyDB(CENTER_DB_LOCAL)

    def find_latest_bank_file(self, directory):
        """가장 최근의 은행 파일 찾기"""
        latest_file = None
        latest_time = 0
        for file in os.listdir(directory):
            file_path = os.path.join(directory, file)
            if not file.endswith('.xls') or not file.startswith('신한은행'):
                continue
            try:
                workbook = xlrd.open_workbook(file_path)
                sheet = workbook.sheet_by_index(0)
                account_no = sheet.cell_value(2, 1)
                if account_no == Account_no['bub']:
                    file_time = os.path.getmtime(file_path)
                    if file_time > latest_time:
                        latest_file, latest_time = file_path, file_time
            except Exception:
                continue
        return latest_file

    def move_file(self, source, destination):
        """파일 이동"""
        try:
            shutil.move(source, destination)
            return True, f"{source} 파일을 {destination}로 이동하였습니다."
        except Exception as e:
            return False, f"파일 이동 실패: {e}"

    def insert_bank_data(self):
        """은행 데이터를 DB에 삽입"""
        try:
            ins_data(self.db)
            return True, "은행 데이터를 성공적으로 삽입하였습니다."
        except Exception as e:
            return False, f"은행 데이터 삽입 실패: {e}"

    def get_pending_bank_data(self):
        """미처리 은행 데이터를 가져오기"""
        sql = """
        SELECT DATE_FORMAT(daytime, '%Y-%m-%d'), contents, FORMAT(deposit, 0), depos_tax 
        FROM bank WHERE depos_tax = %s ORDER BY daytime DESC
        """
        return self.db.run_query(sql, ('F',))

    def get_cash_tax_data(self):
        """현금영수증 미처리 데이터를 가져오기"""
        sql = """
        SELECT id, case_num, input, name, info, bigo, contents 
        FROM mytax WHERE cmplday IS NULL AND cashor = %s
        """
        return self.db.run_query(sql, ('c',))

    def update_cash_tax_status(self, record_id):
        """현금영수증 처리 상태 업데이트"""
        sql = "UPDATE mytax SET cmplday = %s WHERE id = %s"
        self.db.run_query(sql, (datetime.date.today(), record_id))


# View: UI 로직
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QTableWidgetItem
from PyQt5 import uic

class BankView(QMainWindow):
    def __init__(self, ui_file):
        super().__init__()
        uic.loadUi(ui_file, self)
        self.msg = QMessageBox()

    def display_message(self, message):
        """메시지박스 표시"""
        self.msg.setText(message)
        self.msg.exec_()

    def update_table(self, table_widget, data):
        """테이블 위젯 업데이트"""
        table_widget.setRowCount(len(data))
        for i, row in enumerate(data):
            for j, val in enumerate(row):
                table_widget.setItem(i, j, QTableWidgetItem(str(val)))
        table_widget.resizeColumnsToContents()


# Controller: 이벤트 처리 및 데이터 연동
from PyQt5.QtCore import QCoreApplication
import invoice
import time

class BankController:
    def __init__(self, model, view):
        self.model = model
        self.view = view
        self.setup_ui_events()

    def setup_ui_events(self):
        """UI 이벤트 연결"""
        self.view.pushButton.clicked.connect(self.show_cash_tax)
        self.view.pushButton_4.clicked.connect(self.issue_cash_tax)
        self.view.pushButton_3.clicked.connect(self.issue_bill_tax)

    def move_downloaded_bank_files(self):
        """다운로드된 은행 파일 이동"""
        dl_dir = DIR['download']
        bankdata_dir = DIR['bank']

        target_file = self.model.find_latest_bank_file(dl_dir)
        if not target_file:
            self.view.display_message("다운로드 폴더에 법인 은행 파일이 없습니다.")
            return

        today = datetime.date.today()
        destination = os.path.join(bankdata_dir, today.strftime("%y%m") + 'bub.xls')
        success, message = self.model.move_file(target_file, destination)
        self.view.display_message(message)

    def show_cash_tax(self):
        """현금영수증 처리 데이터를 테이블에 표시"""
        make_tax_info(self.model.db)
        data = self.model.get_pending_bank_data()
        self.view.update_table(self.view.tableWidget_2, data)

    def issue_cash_tax(self):
        """현금영수증 처리"""
        cash_data = self.model.get_cash_tax_data()
        if not cash_data:
            self.view.display_message("현금영수증 처리할 데이터가 없습니다.")
            return

        invo = invoice.IssueInvoice()
        self.view.textBrowser.append(f"{datetime.datetime.now()} - 현금영수증 발행 시작")
        self.view.textBrowser.append(invo.login(1))

        for record in cash_data:
            try:
                self.view.textBrowser.append(invo.cash_info(record[6], record[2], record[4]))
                self.model.update_cash_tax_status(record[0])
                self.view.textBrowser.append("DB 업데이트 성공")
            except Exception as e:
                self.view.textBrowser.append(f"현금영수증 처리 실패: {e}")
            time.sleep(2)

        self.view.textBrowser.append(invo.exit())


# Main
if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)

    # Initialize components
    model = BankModel()
    view = BankView(os.path.join(DIR['base'], "pymysql_bank", "tax_Qt.ui"))
    controller = BankController(model, view)

    view.show()
    sys.exit(app.exec_())

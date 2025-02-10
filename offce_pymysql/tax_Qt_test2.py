import os
import sys
import time
import datetime
import calendar
import shutil
import xlrd
import xlwt
from PyQt5.QtWidgets import QMainWindow, QApplication, QMessageBox, QTableWidgetItem
from PyQt5 import uic
from PyQt5.QtCore import Qt
from myconfig import DIR, Account_no, Center_DB_Interner
from mydb import MyDB
from insBankData2Mysql import ins_data, make_tax_info
import invoice

form_class = uic.loadUiType(os.path.join(DIR['base'], "pymysql_bank", "tax_Qt.ui"))[0]


class TaxApp(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.db = MyDB(Center_DB_Interner)

        self.init_ui()
        self.handle_data_movement()
        self.populate_table_widget()

    def init_ui(self):
        self.pushButton.clicked.connect(self.make_cash_tax)
        self.pushButton_4.clicked.connect(self.issue_cash_tax)
        self.pushButton_3.clicked.connect(self.issue_bill_tax)
        self.msg = QMessageBox()

    def handle_data_movement(self):
        try:
            self.move_downloaded_file()
            ins_data(self.db)
        except Exception as e:
            self.textBrowser.append(f"Error: {e}")

    def move_downloaded_file(self):
        bank_dir = DIR['bank']
        download_dir = DIR['download']
        latest_file, latest_time = None, 0

        for file in os.listdir(download_dir):
            full_path = os.path.join(download_dir, file)
            try:
                wbs = xlrd.open_workbook(full_path)
                acc_no = wbs.sheet_by_index(0).cell_value(2, 1)
            except Exception:
                acc_no = None

            if acc_no == Account_no['bub'] and file.endswith('.xls') and file.startswith('신한은행'):
                file_time = os.path.getmtime(full_path)
                if file_time > latest_time:
                    latest_file, latest_time = full_path, file_time

        if latest_file:
            destination = os.path.join(bank_dir, datetime.date.today().strftime("%y%m") + 'bub.xls')
            try:
                shutil.move(latest_file, destination)
                self.textBrowser.append(f"File moved to {destination}")
            except Exception as e:
                self.textBrowser.append(f"Failed to move file: {e}")
        else:
            self.textBrowser.append("No valid file found in download folder.")

    def populate_table_widget(self):
        sql = """SELECT DATE_FORMAT(daytime, '%Y-%m-%d'), contents, format(deposit , 0), depos_tax 
                 FROM bank WHERE depos_tax = 'F' ORDER BY daytime DESC"""
        try:
            # data = self.db.run_query(sql)
            self.populate_table(self.tableWidget, 0) # 임시로 0을 지우고 뒤에 걸 넣어야 됨 data)
        except Exception as e:
            self.textBrowser.append(f"Error loading table data: {e}")

    def populate_table(self, table_widget, data):
        table_widget.setRowCount(len(data))
        for i, row in enumerate(data):
            for j, value in enumerate(row):
                table_widget.setItem(i, j, QTableWidgetItem(str(value)))

    def make_cash_tax(self):
        try:
            make_tax_info(self.db)
            sql = """SELECT DATE_FORMAT(inputday, '%Y-%m-%d'), format(input , 0), contents, cashor 
                     FROM mytax WHERE cmplday IS NULL ORDER BY inputday DESC"""
            data = self.db.run_query(sql)
            self.populate_table(self.tableWidget_2, data)
            self.display_summary()
        except Exception as e:
            self.textBrowser.append(f"Error making cash tax: {e}")

    def display_summary(self):
        mytax_sum, bank_sum, checkup_msg = self.checkup()
        self.textBrowser.append(f"Tax DB Total: {mytax_sum}\nBank Total: {bank_sum}\n{checkup_msg}")

    def checkup(self):
        today = datetime.date.today()
        start, end = today.replace(day=1).strftime("%Y-%m-%d"), today.strftime("%Y-%m-%d")
        sql1 = "SELECT format(SUM(input), 0) FROM mytax WHERE date(inputday) BETWEEN %s AND %s"
        sql2 = "SELECT format(SUM(deposit), 0) FROM bank WHERE date(daytime) BETWEEN %s AND %s"
        
        tax_sum = self.db.run_query(sql1, (start, end))[0][0]
        bank_sum = self.db.run_query(sql2, (start, end))[0][0]

        if tax_sum == bank_sum:
            return tax_sum, bank_sum, "Checkup: Totals match."
        return tax_sum, bank_sum, "Checkup: Totals do not match."

    def issue_cash_tax(self):
        sql = "SELECT id, case_num, input, name, info, bigo, contents FROM mytax WHERE cmplday IS NULL AND cashor = 'c'"
        data = self.db.run_query(sql)

        if not data:
            self.textBrowser.append("No cash receipts to process.")
            return

        invo = invoice.IssueInvoice()
        self.textBrowser.append(invo.login(1))

        for entry in data:
            self.process_cash_tax_entry(entry, invo)

        self.textBrowser.append(invo.exit())

    def process_cash_tax_entry(self, entry, invo):
        try:
            myday = datetime.date.today()
            ut_sql = "UPDATE mytax SET cmplday = %s WHERE id = %s"
            self.db.run_query(ut_sql, (myday, entry[0]))
            self.textBrowser.append(f"Updated mytax for ID {entry[0]}")
        except Exception as e:
            self.textBrowser.append(f"Error updating mytax: {e}")

    def issue_bill_tax(self):
        sql = "SELECT id, inputday, input, name, leader, info, bigo, case_num FROM mytax WHERE cmplday IS NULL AND cashor = 'i'"
        data = self.db.run_query(sql)
        if not data:
            self.textBrowser.append("No invoices to process.")
            return

        file_path = self.create_invoice_file(data)
        self.textBrowser.append(f"Invoice file created: {file_path}")

    def create_invoice_file(self, data):
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('Invoices')

        for i, row in enumerate(data):
            for j, value in enumerate(row):
                sheet.write(i, j, str(value))

        file_name = f"tax{datetime.datetime.now().strftime('%y%m%d')}.xls"
        file_path = os.path.join(DIR['bank'], file_name)
        workbook.save(file_path)
        return file_path


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TaxApp()
    window.show()
    app.exec_()

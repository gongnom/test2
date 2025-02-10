# -*- coding:utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.alert import Alert
# from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
from time import gmtime, strftime
import datetime
import pyautogui
# import cv2
# import pyperclip
import os
from myconfig import DIR
#import autoit
#import selpath
# import requests
# from bs4 import BeautifulSoup
from xml.etree.ElementTree import parse
# from myselenium import get_chromedriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from webdriver_manager.chrome import ChromeDriverManager

def wait(br, text):
    # WebDriverWait(br,30).until(EC.presence_of_element_located((By.XPATH, text)))
    WebDriverWait(br,30).until(EC.element_to_be_clickable((By.XPATH, text)))




def findImage(fileName,  cnt = 10, confidence=0.95, wait=0.1):
    #이미지 발견하여 위치 리턴
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    fileNam = os.path.join(BASE_DIR, fileName)
    for i in range(cnt):
        try: 
            imgpos = pyautogui.locateCenterOnScreen(fileNam, confidence= confidence)
            print("{} found".format(fileNam))
        except: 
            imgpos = None
            print("{} not found".format(fileNam))

        if imgpos != None : 
            print("ok")
            break
        else: 
            print("what problem")
            time.sleep(wait)
    if imgpos == None: 
        return None
    else: 
        return imgpos


class IssueInvoice:


    def __init__(self):

        #get today
        self.dt = datetime.datetime.now()
        self.year = self.dt.strftime("%Y")
        self.month = self.dt.strftime("%m")
        self.day = self.dt.strftime("%d")

        #청구월 (이전달)
        self.first = self.dt.replace(day=1)
        self.last = self.first - datetime.timedelta(days=1)
        self.prevmonth = self.last.strftime("%m")
        self.driver = webdriver.Chrome(ChromeDriverManager().install())

        #드라이버 세팅
        """
        try:
            self.driver = webdriver.Chrome(DIR['chromedriver']) 
        except:
            try:
                print("selenium driver가 없거나 버전이 맞지 않습니다. \n다운로드 중입니다. ")
                print("여기로 가세요 : {}".format(DIR['chromedriver']))
                get_chromedriver() # 현재 동작 안됨
                print("\n다운로드가 완료되었습니다. 잠시만 기다려 주세요.")
                time.sleep(3)
                self.driver = webdriver.Chrome(DIR['chromedriver']) 
            except:
                print("크롬 드라이버를 수동으로 다운로드 받으세요")
        """
        self.driver.maximize_window()
        self.driver.get('https://www.hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index.xml')
        self.main = self.driver.window_handles[0]

        #공인인증서 비밀번호를 넣어주세요
        self.pwd = "partizan10#"
        time.sleep(1)

    def login(self, cash=0):
        #로그인 페이지 이동 //*[@id="anchor22"]
        time.sleep(5)

        self.driver.find_element_by_id('group1543').click()
        # print("login page move. sleep 5")

        time.sleep(5)

        for handle in self.driver.window_handles:
        	self.driver.switch_to_window(self.main)

        #공인인증서 로그인 선택 
        iframe =self.driver.find_element_by_id('txppIframe')
        self.driver.switch_to.frame(iframe)
        # wait(self.driver,'//*[@id="anchor22"' ) 
        self.driver.find_element_by_id('group91882166').click() # group91882124   or  group1655

        iframe = self.driver.find_element_by_id('dscert')
        # print("login page choice. sleep 9")
        time.sleep(1)

        self.driver.switch_to.frame(iframe)

        #공인인증서 선택
        wait(self.driver, '//*[@title="노무법인노동과인권(BizBank)008868020131107188002264"]')
        self.driver.find_element_by_xpath('//*[@title="노무법인노동과인권(BizBank)008868020131107188002264"]') #title="노무법인노동과인권(BizBank)008868020131107188002264"
        self.driver.find_element_by_id('input_cert_pw').send_keys(self.pwd)
        self.driver.find_element_by_id('btn_confirm_iframe').click()
        time.sleep(1)

        #전자세금계산서 발행 화면 
        #self.driver.find_element_by_id('group1300').click()		
        #self.driver.find_element_by_xpath('//*[@id="menuAtag_0104010000"]').click()		
        #self.driver.find_element_by_xpath('//*[@id="menuAtag_0104010100"]').click()		
        #self.driver.find_element_by_id('myMenuImg1').click()	"javascript:fn_topMenuOpen(id);" "menuAtag_0104010100" //*[@id="menuAtag_0104010100"]
        #time.sleep(3)
        #self.driver.find_element_by_id('myMenuQuickA8').click()		
        #time.sleep(5)
        if cash == 0:
            self.goto_voice()
        else:
            self.goto_cash()
        
        dt = datetime.datetime.now()
        return dt.strftime("%Y.%m.%d.%H:%M:%S") + " -  complete login"

    def goto_voice(self):
        # invoice 위치가 자주 바뀜.
        # time.sleep(1)
        # self.driver.find_element_by_id('sub_a_4601010500').click()		
        # self.driver.find_element_by_id('textbox_4601010500').click()		
        time.sleep(5000)
    
    def goto_voice2(self):
        self.driver.find_element_by_id('prcpSrvcMenuA1').click()		
        time.sleep(5)
        
        for handle in self.driver.window_handles:
        	self.driver.switch_to_window(self.main)
        iframe = self.driver.find_element_by_id('txppIframe')
        self.driver.switch_to.frame(iframe)
        self.driver.find_element_by_id('sub_a_3416000000').click()		
        time.sleep(3)


    def customer_info(self, regnumber, checkmonth, bizcondition='', bizitem=''):
        time.sleep(5)
        for handle in self.driver.window_handles:
        	self.driver.switch_to_window(self.main)

        #거래처 조회 클릭
        iframe = self.driver.find_element_by_id('txppIframe')
        print("txppIframe")
        self.driver.switch_to.frame(iframe)
        try:
            self.driver.find_element_by_id('grpBtnDmnrClplcInqrTop').click()
            print("try clicked")
        except:
            self.driver.find_element_by_id('btnDmnrClplcInqrTop').click()
            print("except")

        time.sleep(2)

        #거래처 조회 상세화면 
        iframe =self.driver.find_element_by_id('clplcInqrPopup_iframe')
        self.driver.switch_to.frame(iframe)

        time.sleep(3)
        self.driver.find_element_by_id('group1381').click()
        self.driver.find_element_by_id('edtTxprDscmNoEncCntn').send_keys(regnumber)
        self.driver.find_element_by_id('btnSearch').click()
        self.regnum = regnumber

        time.sleep(1)
        self.driver.find_element_by_id('grdResult_cell_0_0').click()
        self.driver.find_element_by_id('grdResult_cell_0_0').click()

        self.driver.find_element_by_id('btnProcess').click()

        time.sleep(2)
        Alert(self.driver).accept()
        time.sleep(3)
        ###거래처 선택 완료 

        iframe = self.driver.find_element_by_id('txppIframe')
        self.driver.switch_to.frame(iframe)

        #업태변경
        #self.driver.find_element_by_id('edtDmnrBcNmTop').clear()
        #self.driver.find_element_by_id('edtDmnrBcNmTop').send_keys(unicode(bizcondition,'utf-8'))
        #self.driver.find_element_by_id('edtDmnrItmNmTop').clear()
        #self.driver.find_element_by_id('edtDmnrItmNmTop').send_keys(unicode(bizitem,'utf-8'))

        #작성일자
        if checkmonth is False:
            writingdate = self.year+self.prevmonth+self.last
            #writingdate = "20191231" 
            print(writingdate)
            word = self.driver.find_element_by_xpath('//*[@id="calWrtDtTop_input"]')
            actionChains = ActionChains(self.driver)
            actionChains.double_click(word).send_keys(writingdate).perform()
            time.sleep(5)

        dt = datetime.datetime.now()
        return dt.strftime("%Y.%m.%d.%H:%M:%S") + " - complete writing customer information"

    def service_info(self, item, cost, checkmonth, order='0'):

        if checkmonth is True:
            day = self.day
        else:
            day = self.last

        ################--여러개일 수 있음, 0부터 
        #품목일
        box = "genEtxivLsatTop_"+order+"_edtLsatSplDdTop"
        self.driver.find_element_by_id(box).click()
        self.driver.find_element_by_id(box).send_keys(day)

        #품목명
        box = "genEtxivLsatTop_"+order+"_edtLsatNmTop"
        #self.driver.find_element_by_id(box).send_keys(unicode(item,'utf-8'))
        self.driver.find_element_by_id(box).send_keys(item)

        #규격,수량
        box = "genEtxivLsatTop_"+order+"_edtLsatRszeNmTop"
        self.driver.find_element_by_id(box).send_keys("1")

        box = "genEtxivLsatTop_"+order+"_edtLsatQtyTop"
        self.driver.find_element_by_id(box).send_keys("1")

        # 공급가액 계산 시작 --------------------------------------------------------
        #for handle in self.driver.window_handles:
        #	self.driver.switch_to_window(self.main)

        # 계산 클릭
        #iframe = self.driver.find_element_by_id('txppIframe')
        #self.driver.switch_to.frame(iframe)
        box = "genEtxivLsatTop_"+order+"_btnLsatSumClcTop"
        self.driver.find_element_by_id(box).click()
        
        # 공급가액, 세액 자동계산 
        iframe =self.driver.find_element_by_id('calTxAmtSplCftPop_iframe')
        self.driver.switch_to.frame(iframe)
        time.sleep(1)
        
        #self.driver.find_element_by_id('edtSumAmt').click()
        self.driver.find_element_by_id('edtSumAmt').send_keys(cost)
        self.driver.find_element_by_id('btnLsatSumTop').click()
        self.driver.find_element_by_id('btnConfirm').click()
        time.sleep(1)
        
        iframe = self.driver.find_element_by_id('txppIframe')
        self.driver.switch_to.frame(iframe)

        # 단가로 할때 
        #box = "genEtxivLsatTop_"+order+"_edtLsatUtprcTop"
        #self.driver.find_element_by_id(box).send_keys(cost)
        #공급가액, 세액은 클릭만 하면 자동계산
        #box = "genEtxivLsatTop_"+order+"_edtLsatSplCftTop"
        #self.driver.find_element_by_id(box).click()

        #공급가액 계산 끝-----------------------------------------------------------

        # 청구, 영수 체크 rdoRecApeClCdTop_input_1
        # if self.regnum is not "5848200178":     # 서대문센터면 청구에 체크(가만히 놔둠)
        self.driver.find_element_by_id("rdoRecApeClCdTop_input_1").click()

        print("item : " + item)
        dt = datetime.datetime.now()
        return dt.strftime("%Y.%m.%d.%H:%M:%S") + " - complete writing information"

    def issue(self):

        #발급
        self.driver.find_element_by_id('btnIsn').click()
        time.sleep(3)

        #email 없다는 창이 뜰때
        try: 
            Alert(self.driver).accept()
            #alert = self.driver.switch_to_alert()
            #alert.accept()
        except:
            pass

        #iframe 전환, 전자세금계산서 발급 확인 팝업
        time.sleep(3)
        self.iframe = self.driver.find_element_by_id('UTEETZZA89_iframe')
        self.driver.switch_to.frame(self.iframe)

        #발급취소
        #self.driver.find_element_by_id('btnClose').click()

        #발급하기
        self.driver.find_element_by_id('trigger20').click()

        time.sleep(5)

        #서명취소
        ##################################################################
        #ControlFocus("인증서 선택창", " ", "Edit2")
        #ControlSetText("인증서 선택창"," ", "Edit2", "인증서 비밀번호")
        #ControlClick("인증서 선택창", " ", "Button12")
        ##################################################################		
        #autoit.run("C:\\work\\invoice\\test_sign.exe")

        #인증서 서명
        try:
            #imgpos = pyautogui.locateCenterOnScreen("cert.jpg", confidence= 0.8)
            #imgpos2 = pyautogui.locateCenterOnScreen("cert_ok.jpg", confidence= 0.8)
            pyautogui.click(findImage('cert1.png', confidence=.8))
            #pyautogui.click(findImage('cert.jpg', confidence=.8))
        except:
            pyautogui.click(findImage('cert.jpg', confidence=.8))
            #pyautogui.click(findImage('cert.png', confidence=.8))

        time.sleep(1)
        try:
            pyautogui.typewrite(self.pwd, interval=0.1)
        except:
            pass
        
        try:
            pyautogui.click(findImage('cert_ok1.png', confidence=.8))
            #pyautogui.click(findImage('cert_ok.jpg', confidence=.8))
        except:
            pyautogui.click(findImage('cert_ok.jpg', confidence=.8))
            #pyautogui.click(findImage('cert_ok.png', confidence=.8))


        ##################################################################
        #ControlFocus("인증서 선택창", " ", "Edit2")
        #ControlSetText("인증서 선택창"," ", "Edit2", "인증서 비밀번호")
        #ControlClick("인증서 선택창", " ", "Button11")
        ##################################################################
        #autoit.run("C:\\work\\invoice\\invoice_sign.exe")

        dt = datetime.datetime.now()
        return dt.strftime("%Y.%m.%d.%H:%M:%S") + " - complete issuing invoice"		
    
    def cont_issue(self):
        # iframe id = "isnCmplPopup_iframe"
        # input button id ="trigger20"
        for handle in self.driver.window_handles: 
            print("handle: \t") 
            print(handle) 
            self.driver.switch_to_window(self.main)
        
        #iframe =self.driver.find_element_by_id('clplcInqrPopup_iframe')
        #self.driver.switch_to.window(self.driver.window_handles[1])

        #time.sleep(3)
        #self.driver.find_element_by_id('group1381').click()
        try:
            time.sleep(1)
            iframe = self.driver.find_element_by_id('isnCmplPopup_iframe')
            self.driver.switch_to.frame(iframe)
            self.driver.find_element_by_id('trigger20').click()
        except:
            time.sleep(1)
        
        dt = datetime.datetime.now()
        return dt.strftime("%Y.%m.%d.%H:%M:%S") + " - continue next invoice."		

    def goto_cash(self):
        # 얘도 한번씩 바뀜. prcpSrvcMenuA3, prcpSrvcMenuA5 group8861604543 group8861604545 sub_a_0105080000
       

        time.sleep(5)
        
        for handle in self.driver.window_handles:
        	self.driver.switch_to_window(self.main)
        iframe = self.driver.find_element_by_id('txppIframe')
        self.driver.switch_to.frame(iframe)
        # 현금영수증 건별 발급
        self.driver.find_element_by_id('sub_a_4606020100').click()		
        time.sleep(2)
        # self.driver.find_element_by_id('sub_a_3301000000').click()		
        # time.sleep(3)

        # 승인거래 발급
        # self.driver.find_element_by_id('group1532').click()		
        # time.sleep(3)

    def cash_info(self, name, total, num): 
        for handle in self.driver.window_handles:
        	self.driver.switch_to_window(self.main)
        iframe = self.driver.find_element_by_id('txppIframe')
        self.driver.switch_to.frame(iframe)
        time.sleep(8)

        if num == '1234':
            self.driver.find_element_by_id('radio8_input_0').click()		
        else :
            #self.driver.find_element_by_id('radio8_input_1').click()		
            self.driver.find_element_by_id('spstCnfrNoEncCntn').send_keys(num)
        
        time.sleep(1)
        self.driver.find_element_by_id('trsAmt').send_keys(total)
        time.sleep(1)
        # 비고
        self.driver.find_element_by_id('cshptIsnMmoCntn').send_keys(name)
        time.sleep(1)
        self.driver.find_element_by_id('trigger4').click()		

        try:
            time.sleep(3) 
            Alert(self.driver).accept()
        except:
            pass
        
        try:
            time.sleep(3) 
            Alert(self.driver).accept()
        except:
            pass
        
        time.sleep(3) 
        self.driver.find_element_by_id('trigger27').click()		
        #trigger27
        time.sleep(3) 
        
        dt = datetime.datetime.now()
        return dt.strftime("%Y.%m.%d.%H:%M:%S") +name+ " - complete writing cash information"
        
   
    def exit(self):
        self.driver.quit()

        dt = datetime.datetime.now()
        return dt.strftime("%Y.%m.%d.%H:%M:%S") + " - close all window"
        

if __name__ == "__main__":
    a = IssueInvoice()
    a.login()
    # test()
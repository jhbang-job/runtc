######################################
#버전 : 0.3
#사용방법 : c:\> python runtc.py [엑셀경로]
#          GUI 실행
######################################

import sys
import pyautogui
import pyperclip
import time
import xlwings as xw
from PyQt5.QtWidgets import *
from PIL import ImageGrab 
from wand.image import Image # 설치 필요 : https://imagemagick.org/script/download.php#windows
    
######################################
# GUI
######################################
class MyWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setupUI()
        self.result = None

    def setupUI(self):
        self.setGeometry(800, 200, 300, 300)
        self.setWindowTitle("runtc v0.3")

        self.pushButton = QPushButton("불러오기")
        self.pushButton.clicked.connect(self.pushButtonClicked)
        self.label = QLabel()

        self.btn1 = QPushButton("실행", self)
        self.btn1.move(11, 40)
        self.btn1.clicked.connect(self.btn1_clicked)

        layout = QVBoxLayout()
        layout.addWidget(self.pushButton)
        layout.addWidget(self.btn1)
        layout.addWidget(self.label)

        self.setLayout(layout)

    def pushButtonClicked(self):
        fname = QFileDialog.getOpenFileName(self)
        self.label.setText(fname[0])
        self.result = fname[0]
        return self.result


    def btn1_clicked(self, result):
        
        txt = main(self.result)
        QMessageBox.about(self, "message", txt)
        

######################################
# pyautogui
######################################


def 클릭(**kwagrs):
    try:
        kwagrs['정확도']
    except:
        kwagrs['정확도'] = 0.7
    
    five_btn = pyautogui.locateOnScreen(kwagrs['이미지'], confidence=kwagrs['정확도'])
    

    #이미지 영역의 가운데 위치 얻기
    five_btn = pyautogui.locateOnScreen(kwagrs['이미지'], confidence=kwagrs['정확도'])
    center = pyautogui.center(five_btn)
    
    
    #클릭하기
    #center = pyautogui.locateCenterOnScreen(kwagrs['이미지'], confidence=kwagrs['정확도'])
    print("1.클릭함수", kwagrs['이미지'], kwagrs['정확도'])
    print("2.찾아보자 아이콘", five_btn)
    print("3.찾아보자 아이콘센터 ", center)
    pyautogui.click(center)


def 우클릭(**kwagrs):
    try:
        kwagrs['정확도']
    except:
        kwagrs['정확도'] = 0.7
    
    five_btn = pyautogui.locateOnScreen(kwagrs['이미지'], confidence=kwagrs['정확도'])
    print("찾아보자 아이콘", five_btn)

    #이미지 영역의 가운데 위치 얻기
    five_btn = pyautogui.locateOnScreen(kwagrs['이미지'], confidence=kwagrs['정확도'])
    center = pyautogui.center(five_btn)
    
    
    print(kwagrs['이미지'], kwagrs['정확도'])
    #클릭하기
    center = pyautogui.locateCenterOnScreen(kwagrs['이미지'], confidence=kwagrs['정확도'])
    print("찾아보자 아이콘센터 ", center)
    pyautogui.click(center, button='right')
    screenshot()


#더블클릭(이미지=이미지명, 정확도=0.7)
def 더블클릭(**kwagrs):

    try:
        kwagrs['정확도']
    except:
        kwagrs['정확도'] = 0.7
    
    five_btn = pyautogui.locateOnScreen(kwagrs['이미지'], confidence=kwagrs['정확도'])
    #print(five_btn)

    #이미지 영역의 가운데 위치 얻기
    five_btn = pyautogui.locateOnScreen(kwagrs['이미지'], confidence=kwagrs['정확도'])
    center = pyautogui.center(five_btn)
    #print(center)
    
    #클릭하기
    center = pyautogui.locateCenterOnScreen(kwagrs['이미지'], confidence=kwagrs['정확도'])
    
    print("1.클릭함수", kwagrs['이미지'], kwagrs['정확도'])
    print("2.찾아보자 아이콘", five_btn)
    print("3.찾아보자 아이콘센터 ", center)
    pyautogui.doubleClick(center)

def 키보드누르기(변수):
    pyautogui.press(변수)

def 입력(변수):
    pyautogui.typewrite(변수)

def 한글입력(변수):
    pyperclip.copy(변수)
    pyautogui.hotkey("ctrl", "v")

def 핫키(*agrs):
    pyautogui.hotkey(*agrs)

def 시간지연(초):
    time.sleep(초)

######################################
# 엑셀 제어
######################################

#file = "d:/runtc/tmp/tcrun_test.xlsx"
#wb = xw.Book(file)
wb = None

def image읽기(이미지):
    global wb
    sht = wb.sheets['image']
    a = sht.pictures(이미지)
    a.api.Copy()

    Image(filename='clipboard:').save(filename='test.png')
    img = Image(filename='clipboard:')
    
    return img

def tc읽기():
    global wb
    sht = wb.sheets['tc']
    endLine = wb.sheets['tc'].used_range.address.split('$')[4]
    Line = sht.api.Cells.Find('tcrun')
    address = sht.range((Line.Row, Line.Column)).address.split('$')[1]
    used_range = str(address) + str(Line.Row+1) + ":" + str(address) + str(endLine)
    tcRanage = sht.range(used_range)
    

    return tcRanage


def 명령_수행(명령, 매개변수):
    
    if "!클릭" == 명령 :
        if len(매개변수) == 2:
            명령_이미지, 정확도 = 매개변수
        elif len(매개변수) == 1:
            명령_이미지 = 매개변수[0]
            정확도 = 0.7
        else:
            print("입력값오류")
            pass
    
        if ">" in 명령_이미지:
            for i in 명령_이미지.split(' > '):
                읽어온_이미지 = image읽기(i[1:])
                클릭(이미지=읽어온_이미지, 정확도=정확도)
                time.sleep(0.5)
        else:
            읽어온_이미지 = image읽기(명령_이미지[1:])
            클릭(이미지=읽어온_이미지, 정확도=정확도)
        
    
    elif "!더블클릭" == 명령:
        if len(매개변수) == 2:
            명령_이미지, 정확도 = 매개변수
        elif len(매개변수) == 1:
            명령_이미지 = 매개변수[0]
            정확도 = 0.7
        else:
            print("입력값오류")
            pass
        
        if ">" in 명령_이미지:
            for i in 명령_이미지.split(' > '):
                읽어온_이미지 = image읽기(i[1:])
                클릭(이미지=읽어온_이미지, 정확도=정확도)
                time.sleep(0.5)
        else:
            읽어온_이미지 = image읽기(명령_이미지[1:])
            더블클릭(이미지=읽어온_이미지, 정확도=정확도)

    if "!우클릭" == 명령 :
        if len(매개변수) == 2:
            명령_이미지, 정확도 = 매개변수
        elif len(매개변수) == 1:
            명령_이미지 = 매개변수[0]
            정확도 = 0.7
        else:
            print("입력값오류")
            pass
    
        if ">" in 명령_이미지:
            for i in 명령_이미지.split(' > '):
                읽어온_이미지 = image읽기(i[1:])
                우클릭(이미지=읽어온_이미지, 정확도=정확도)
                time.sleep(0.5)
        else:
            읽어온_이미지 = image읽기(명령_이미지[1:])
            우클릭(이미지=읽어온_이미지, 정확도=정확도)


    elif "!키보드누르기" in 명령:
        키보드누르기(*매개변수)
    
    elif "!입력" in 명령:
        입력(*매개변수)
    
    elif "!핫키" in 명령:
        핫키(*매개변수)
        
    elif "!시간지연" in 명령:
        시간지연(float(*매개변수))


def TC문법_수행(tc):
    tc한줄분리 = tc.replace(';','\n').split('\n')
    for i in tc한줄분리:
        if "!" in i:
            한줄 = i.strip().replace('"','').replace('\'','').replace('(','@@').replace(')','@@').split('@@')
            명령 = 한줄[0]
            매개변수 = 한줄[1:-1][0].split(', ')
            명령_수행(명령, 매개변수)
            print(매개변수)
            
            
                
                
            
            
            
######################################
# 테스트 수행
######################################
def main(file):
    global wb
    if file != None and "xlsx" in file :
        wb = xw.Book(file)
        tc = tc읽기()
        tc = tc.value
        for i in tc:
            try:
                TC문법_수행(i)
            except:
                messge = "테스트 종료"
                print(messge)
                return messge
    
    else:
        messge = "xlsx 파일이 아닙니다."
        print(messge)
        return messge
        #sys.exit(1)
    

def gui():
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()

def run():
    '''
    #시작버튼 = image읽기('시작버튼')
    #더블클릭(이미지=시작버튼)
    매개변수 = ['win', 'd']
    핫키(*매개변수)
    핫키('win', 'r')
    시간지연(2)
    입력('chrome https://www.naver.com')
    키보드누르기("'enter'")
    시간지연(2)
    핫키("'win', 'up'")
    '''
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()
    
    
if __name__=="__main__":
    try:
        file = sys.argv[1]
        
    except:
        file = None
        #file = "D:/runtc/tc/runtc_sample.xlsx"
        
    #main(file)
    if file == None:
        gui()
    else:
        main(file)
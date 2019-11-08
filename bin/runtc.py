######################################
#버전 : 0.2
#사용방법 : c:\> python runtc.py [엑셀경로]
#          GUI 실행
######################################

import sys
import pyautogui  
import time
import xlwings as xw
from PyQt5.QtWidgets import *

	
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
		self.setWindowTitle("runtc v0.2")

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
	#print(five_btn)

	#이미지 영역의 가운데 위치 얻기
	five_btn = pyautogui.locateOnScreen(kwagrs['이미지'], confidence=kwagrs['정확도'])
	center = pyautogui.center(five_btn)
	#print(center)
	
	#클릭하기
	center = pyautogui.locateCenterOnScreen(kwagrs['이미지'], confidence=kwagrs['정확도'])
	pyautogui.click(center)

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
	pyautogui.doubleClick(center)

def 키보드누르기(변수):
	pyautogui.press(변수)

def 입력(변수):
	pyautogui.typewrite(변수)

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
	img_b = a.api.Copy()

	from PIL import ImageGrab 
	pic = ImageGrab.grabclipboard()
	return pic

def tc읽기():
	global wb
	sht = wb.sheets['tc']
	endLine = wb.sheets['tc'].used_range.address.split('$')[4]
	Line = sht.api.Cells.Find('tcrun')
	address = sht.range((Line.Row, Line.Column)).address.split('$')[1]
	used_range = str(address) + str(Line.Row+1) + ":" + str(address) + str(endLine)
	tcRanage = sht.range(used_range)
	

	return tcRanage


def 이미지_처리(파라미터):
	정확도 = 0.7
	if 파라미터[0] == "#":
		try:
			이미지,정확도 = 파라미터.split(',')
		except:
			이미지 = 파라미터
		
		파라미터 = image읽기(이미지[1:])
		
	결과 = [파라미터, 정확도]
	
	return 결과


def 함수_수행(함수명, 파라미터):
	if "!클릭" in 함수명 :
		if len(파라미터) == 2:
			파라미터, 정확도 = 파라미터
		elif len(파라미터) == 1:
			파라미터 = 파라미터[0]
			정확도 = 0.7
		else:
			print("입력값오류")
			pass
		
		if ">" in 파라미터:
			for i in 파라미터.split(' > '):
				파라미터 = image읽기(i[1:])
				클릭(이미지=파라미터, 정확도=정확도)
				time.sleep(0.5)
		else:
			파라미터 = image읽기(파라미터[1:])
			클릭(이미지=파라미터, 정확도=정확도)
			
	
	elif "!더블클릭" in 함수명:
		if len(파라미터) == 2:
			파라미터, 정확도 = 파라미터
		elif len(파라미터) == 1:
			파라미터 = 파라미터[0]
			정확도 = 0.7
		else:
			print("입력값오류")
			pass
		
		if ">" in 파라미터:
			for i in 파라미터.split(' > '):
				파라미터 = image읽기(i[1:])
				더블클릭(이미지=파라미터, 정확도=정확도)
		else:
			파라미터 = image읽기(파라미터[1:])
			더블클릭(이미지=파라미터, 정확도=정확도)

	elif "!키보드누르기" in 함수명:
		키보드누르기(*파라미터)
	
	elif "!입력" in 함수명:
		입력(*파라미터)
	
	elif "!핫키" in 함수명:
		핫키(*파라미터)
		
	elif "!시간지연" in 함수명:
		시간지연(float(*파라미터))


def TC문법_수행(tc):
	tc한줄분리 = tc.replace(';','\n').split('\n')
	for i in tc한줄분리:
		if "!" in i:
			한줄 = i.strip().replace('"','').replace('\'','').replace('(','@@').replace(')','@@').split('@@')
			함수명 = 한줄[0]
			파라미터 = 한줄[1:-1][0].split(', ')
			함수_수행(함수명, 파라미터)
			
				
				
			
			
			
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
	파라미터 = ['win', 'd']
	핫키(*파라미터)
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
	

from cryptography.fernet import Fernet
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtCore import QDate
import datetime as dt
import sys
import os
from posixpath import split

# MS)820,20220409/ALIGN)820,20220409/VPN)12:12:12:12:12:12
# MS)DISABLE/ALIGN)DISABLE/VPN)DISABLE

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UI_PATH = "code_maker.ui"

class MainDialog(QDialog) :
    def __init__(self) :
        QDialog.__init__(self, None)
        uic.loadUi(os.path.join(BASE_DIR, UI_PATH), self)
        
        # 클래스 속성
        ## 대칭키
        self.key_code = "0rKTqtZ4GQJlecUE8zRfU1B-metwX61_2Iz6B66B5eo="
        self.key_code = self.key_code.encode('utf-8')
        self.fernet = Fernet(self.key_code)
        ## 오늘 날짜 
        self.today = dt.datetime.today()
        ## 종료 날짜
        # self.excel_conv_encry_expiration_dateEdit.setDate(QDate(self.today))
        # 라디오 버튼 체크
        self.excel_conv_able_radioButton.setChecked(True)

        # 동적 매서드
        self.encry_btn.clicked.connect(self.encryptography)                 # 암호화 시작 버튼(암호화 함수 실행)
        self.decry_btn.clicked.connect(self.decryptography)                 # 복호화 시작 버튼(복호화 함수 실행)
        self.excel_conv_able_radioButton.toggled.connect(self.able)         # 자동 VPN ON/OFF 라디오 버튼(활성화 함수 실행)

        self.excel_conv_encry_expiration_dateEdit.dateChanged.connect(self.service_period_calculator)  # 날짜 상하 조절버튼(암호화기 그룹박스 서비스기간 출력 함수 실행)

    # 활성화 함수
    def able(self) :
        if self.excel_conv_able_radioButton.isChecked() == True :
            self.excel_conv_mac_address_line_edit.setEnabled(True)
            self.excel_conv_mac_address_line_edit.setText('')
        elif self.excel_conv_disable_radioButton.isChecked() == True :
            self.excel_conv_mac_address_line_edit.setEnabled(False)
            self.excel_conv_mac_address_line_edit.setText('open')

    # 암호화 함수
    def encryptography(self) :
        self.service_period_calculator()                                    # 암호화기 그룹박스 서비스기간 출력 함수 실행.

        # 자동 창 정렬 암호화활 텍스트 만들기
        if self.excel_conv_able_radioButton.isChecked() == True :
            excel_conv_mac_address = str(self.excel_conv_mac_address_line_edit.text())            # int로 되어있는 TMG번호를 line_edit에서 str로 변환하여 tmg 번호를 가져옴.
            if excel_conv_mac_address == "" :
                self.encry_license_text_line_edit.setText('mac 주소를 입력해 주세요.')
                self.encry_text_status.setText('mac 주소를 입력해 주세요.')
                return 0
            excel_conv_encry_expiration = str(self.excel_conv_encry_expiration_dateEdit.text())       # 종료날짜 date_edit로 부터 str로 변환하여 종료날짜를 가져옴.
            excel_conv_encry_text = f'{excel_conv_mac_address},{excel_conv_encry_expiration}'

        elif self.excel_conv_disable_radioButton.isChecked() == True :
            excel_conv_mac_address = str(self.excel_conv_mac_address_line_edit.text())            # int로 되어있는 TMG번호를 line_edit에서 str로 변환하여 tmg 번호를 가져옴.
            excel_conv_encry_expiration = str(self.excel_conv_encry_expiration_dateEdit.text())       # 종료날짜 date_edit로 부터 str로 변환하여 종료날짜를 가져옴.
            excel_conv_encry_text = f'{excel_conv_mac_address},{excel_conv_encry_expiration}'

        # 암호화 텍스트 만들기 및 암호화 진행
        encry_text = f'{excel_conv_encry_text}'                             # TMG번호와 종료날짜를 조합하여 암호화 할 문자열 생성. string 형태
        self.encry_text_status.setText(encry_text)                          # 암호화할 문자열을 label에 출력. 

        license = self.fernet.encrypt(encry_text.encode('utf-8'))           # str로 되어있는 암호화할 문자열을 인코딩하여 binary형태로 만든 후 암호화하고 license에 저장
        license = license.decode('utf-8')                                   # binary 형태로 저장되어있는 license를 디코딩하여 str 형태로 변형.

        self.encry_license_text_line_edit.setText(license)                  # license를 line_edit에 출력.

    # 복호화 함수
    def decryptography(self) :                                              
        license = self.decry_license_text_line_edit.text()                  # line_edit에서 추출한 텍스트를 licence변수에 입력.
        if license == "" :
            self.decry_text_status.setText('라이센스키를 입력해주세요.')
            return 0

        try :
            decry_text = self.fernet.decrypt(license.encode('utf-8'))           # license를 인코딩하여 바이너리형태로 만들고, 복호화 하여 decry_text에 저장
            decry_text = decry_text.decode('utf-8')                             # decry_text를 디코딩하여 str형태로 변형.
            
            excel_conv_decry_data_list = decry_text.split(',')
            excel_conv_decry_mac = excel_conv_decry_data_list[0]
            excel_conv_decry_expiration = excel_conv_decry_data_list[1]
            excel_conv_decry_expiration_date_list = excel_conv_decry_expiration.split('-')            # mapping_decry_expiration을 -로 나눠 연,월,일로 나눔.
            excel_conv_decry_expiration_year = int(excel_conv_decry_expiration_date_list[0])          # 연 정보가 입력된 str 데이터를 int로 변형.
            excel_conv_decry_expiration_monce = int(excel_conv_decry_expiration_date_list[1])         # 월 정보가 입력된 str 데이터를 int로 변형.
            excel_conv_decry_expiration_day = int(excel_conv_decry_expiration_date_list[2])         # 일 정보가 입력된 str 데이터를 int로 변형.
            excel_conv_decry_expiration_dt = dt.datetime(excel_conv_decry_expiration_year,excel_conv_decry_expiration_monce,excel_conv_decry_expiration_day)    # 날짜 계산을 위해 연,월,일 정보를 합하여 datetime 데이터로 변형함.
            excel_conv_decry_service_period = excel_conv_decry_expiration_dt - self.today             # [만기일 - 현재일 = 서비스 기간] 식을 만들어 줌.
            self.excel_conv_decry_mac_status.setText(excel_conv_decry_mac)
            self.excel_conv_decry_expiration_status.setText(excel_conv_decry_expiration)
            self.excel_conv_decry_service_period_status.setText(str(int(excel_conv_decry_service_period.days)+1))    # label에 서비스 기간 출력.
            self.decry_text_status.setText(decry_text)                          # label에 복호화된 문자열 출력.

        except :
            self.decry_text_status.setText('복호화에 실패하였습니다.')

        # 암호화기 그룹박스 서비스기간 출력 함수
    def service_period_calculator(self) :
        # 암호화기 종료날짜에서 날짜 데이터를 가져옴
        excel_conv_encry_expiration_date = self.excel_conv_encry_expiration_dateEdit.text()

        # 날짜 데이터를 '-'로 나눠서 연,월,일 데이터를 리스트로 분리
        excel_conv_encry_expiration_date_list = excel_conv_encry_expiration_date.split('-')

        # 연 정보가 입력된 str 데이터를 int로 변형
        excel_conv_encry_expiration_year = int(excel_conv_encry_expiration_date_list[0])

        # 월 정보가 입력된 str 데이터를 int로 변헝
        excel_conv_encry_expiration_monce = int(excel_conv_encry_expiration_date_list[1])

        # 일 정보가 입력된 str 데이터를 int로 변형
        excel_conv_encry_expiration_day = int(excel_conv_encry_expiration_date_list[2])

        # 날짜 계산을 위해 연,월,일 정보를 합하여 datetime 데이터로 변형함.
        excel_conv_encry_expiration = dt.datetime(excel_conv_encry_expiration_year,excel_conv_encry_expiration_monce,excel_conv_encry_expiration_day)

        # [만기일 - 현재일 = 서비스 기간] 식을 만들어 줌.
        excel_conv_encry_service_period = excel_conv_encry_expiration - self.today

        # 서비스 기간 데이터들 서비스기간 label에 출력
        self.excel_conv_encry_service_period_status.setText(str(int(excel_conv_encry_service_period.days)+1))

QApplication.setStyle("fusion")
app = QApplication(sys.argv)
main_dialog = MainDialog()
main_dialog.show()

sys.exit(app.exec_())
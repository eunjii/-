######인터파크 쇼핑케어######
import time
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

# 현재 크롬 창에서 진행
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
chrome_driver = "C:/Users/Owner/AppData/Local/Programs/Python/chromedriver.exe"
driver = webdriver.Chrome(chrome_driver, options=chrome_options)

#########################################################문자 인증 함수 정의############################################################################
# 얼럿 처리
def alert_ok():
    confirm = Alert(driver)
    confirm.accept()
    print("얼럿 처리 완료")

# 문자 수신 여부 확인
def message_check():
    i = 0
    number0 = driver.find_element_by_xpath('/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-main-nav/mws-conversations-list/nav/div[1]/mws-conversation-list-item[1]').text  # 발신번호 출력
    number = number0[0:9]  # 번호만 출력
    print(number)
    if number == '1599-1874':
        print("새 문자 수신 완료")
        extract_message()
    else:
        while i < 4: # 횟수 변경
            print("새 문자 수신 미완료")
            driver.refresh()
            print("새로고침")
            time.sleep(5) # 시간 변경
            number2 = driver.find_element_by_xpath('/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-main-nav/mws-conversations-list/nav/div[1]/mws-conversation-list-item[1]').text  # 발신번호 출력
            number3 = number2[0:9]  # 번호만 출력
            if number3 == '1599-1874':
                print("새로고침 후 문자 수신 완료")
                extract_message()
                return
            else:
                i += 1
            continue
        re_message()  #재발송 함수 호출

# 최신 문자 출력
str = "sms_number"
def extract_message():
    global sms_number
    recent_sms = driver.find_element_by_xpath('/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-main-nav/mws-conversations-list/nav/div[1]/mws-conversation-list-item[1]').text
    sms_number = recent_sms[29:35]
    print("최신 문자 : "+sms_number)
    driver.close()
    driver.switch_to.window(tabs[-1])
    time.sleep(2)
    return

# 인증번호 재발송
def re_message():
    driver.switch_to.window(tabs[-1])
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="retryRequst"]').send_keys(Keys.ENTER)
    time.sleep(2)
    print("재발송 성공")
    alert_ok()
    time.sleep(5)
    driver.switch_to.window(tabs[0])
    time.sleep(2)
    message_check()  # 문자 수신 여부 확인

# 인증 후 다음 스탭으로
def auth_mobile():
    driver.switch_to.window(tabs[-1])

    # 인증 번호 넣기
    driver.find_element_by_name('confirm').send_keys(sms_number)
    time.sleep(3)

    # 약관 모두 동의
    auth_agree_check = driver.find_element_by_xpath('//*[@id="chkboxAll"]')
    driver.execute_script("arguments[0].click();", auth_agree_check)
    time.sleep(3)

    # 다음 스탭 이동
    driver.find_element_by_xpath('//*[@id="registerBtn"]').send_keys(Keys.ENTER)
    print("메일 인증 페이지로 이동")
    # 오입력 처리
    try:
        WebDriverWait(driver, 5).until(EC.alert_is_present(),
                                       'Timed out waiting for PA creation ' +
                                       'confirmation popup to appear.')

        alert = driver.switch_to.alert
        alert.accept()
        print("인증번호 재발송 필요")
        re_message()
    except:
        print("pass")

################################################################문자 인증 실행########################################################################
# 현재 탭에서 메시지 앱 진입
driver.get("https://messages.google.com/web/conversations/52")
time.sleep(3)

# 새 탭 추가
driver.execute_script("window.open();")

# 새 탭에서 인터파크 홈페이지 접속
tabs = driver.window_handles
driver.switch_to.window(tabs[-1])

driver.get("https://interparkcare.co.kr/")
time.sleep(2)

# 인터파크 가입 페이지 접속
driver.find_element_by_xpath('//*[@id="join"]').send_keys(Keys.ENTER)
time.sleep(3)

# 휴대폰 인증 요청
driver.find_element_by_name('tel').send_keys('99658939')
driver.find_element_by_xpath('//*[@id="container"]/section/article/section[2]/div[1]/span/a').send_keys(Keys.ENTER)
time.sleep(15)

# 처음 탭으로 이동
driver.switch_to.window(tabs[0])

# beautifulsoup 라이브러리 선언
html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')

message_check()  # 문자 수신 여부 체크
time.sleep(3)
auth_mobile()  # 인증 후 다음 스탭
time.sleep(5)

#########################################################메일 인증 함수 정의##########################################################################
# 새 메일 수신 여부 체크
def nonread(xpath):
    try:
        driver.find_element_by_xpath(xpath)
        return True
    except:
        return False

# 새 인증번호 출력
str = "extract_code"
def extract_mcode():
    global extract_code
    code = driver.find_element_by_class_name("y2").text
    extract_code = code[8:14]
    print("새 인증번호 : "+extract_code)
    driver.find_element_by_class_name("xS").click()  # 메일 읽음 처리
    time.sleep(3)
    driver.close()
    driver.switch_to.window(tabs[-1])
    return

# 메일 재발송
def re_mail():
    tabs = driver.window_handles
    driver.switch_to.window(tabs[1])
    time.sleep(3)
    driver.find_element_by_xpath('//*[@id="emailAuthRetryBtn"]').send_keys(Keys.ENTER)
    time.sleep(4)
    driver.find_element_by_xpath('//*[@id="check_code"]/div/div/span/button').send_keys(Keys.ENTER)
    time.sleep(4)
    confirm = Alert(driver)
    confirm.accept()
    print("메일 재발송 완료")
    driver.switch_to.window(tabs[-1])
    print("메일 기다리는 중..1분")
    time.sleep(60)  #메일 발송 평균 시간
    em_check()  # 메일 수신 체크

# 메일 발송 및 페이지 이동
def em_main():
    # 이메일 발송
    driver.find_element_by_name('joinEmail').send_keys('dmawl1592@gmail.com')  # 이메일 입력
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="sendAuthEmail"]').send_keys(Keys.ENTER)  # 이메일 발송
    time.sleep(2)

    # 이메일 발송 확인 얼럿 처리
    alert_ok()

    # 새 탭 추가
    driver.execute_script("window.open();")

    # 구글 메일 진입
    tabs = driver.window_handles
    driver.switch_to.window(tabs[-1])
    print("메일 기다리는 중..1분")
    time.sleep(60)  # 시간 바꾸기

    driver.get('https://mail.google.com/mail/u/0/?tab=rm&ogbl#advanced-search/query=%5B%EC%9D%B8%ED%84%B0%ED%8C%8C%ED%81%AC+%EC%87%BC%ED%95%91%EC%BC%80%EC%96%B4%5D+%ED%9A%8C%EC%9B%90%EA%B0%80%EC%9E%85%EC%9D%84+%EC%9C%84%ED%95%9C+%EC%9D%B8%EC%A6%9D%EC%BD%94%EB%93%9C%EB%A5%BC+%EC%A0%84%EB%8B%AC%EB%93%9C%EB%A6%BD%EB%8B%88%EB%8B%A4.&isrefinement=true')
    time.sleep(3)

# 새 메일 수신 여부 확인
def em_check():
    i = 0
    if nonread('//*[@id=":4"]/div/div[1]/div/div[1]/div/div[5]'):  # 새 메일 있으면 다음 스탭
        print("새 메일 수신 완료")
        extract_mcode()
    else:  # 새 메일 기다리기
        while i <= 3:
            print("새 메일 수신 미완료")
            driver.refresh()
            time.sleep(10)
            if nonread('//*[@id=":4"]/div/div[1]/div/div[1]/div/div[5]'):
                print("새로고침 후 메일 수신 완료")
                extract_mcode()
                return
            else:
                i += 1
                continue
        re_mail()

# 무료 가입 완료
def free_join():
    driver.find_element_by_name('email_code').send_keys(extract_code) # 인증번호 입력
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="emailAuthArea"]/span/a').send_keys(Keys.ENTER)  # 인증확인 버튼 탭
    time.sleep(2)
    alert_ok()
    print("이메일 인증 완료")
    time.sleep(2)
    driver.find_element_by_name('pass1').send_keys('123456a!')  # 비밀번호 입력
    time.sleep(1)
    driver.find_element_by_name('pass2').send_keys('123456a!')  # 비밀번호 재입력
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="registerMemberBtn"]').send_keys(Keys.ENTER)  # 가입
    print("무료 가입 성공")

################################################################메일 인증 실행######################################################################
em_main()
em_check()
time.sleep(4)
free_join()

#############################################################결제 함수 정의##########################################################################
# 결제
def pay():
    time.sleep(2)
    # 약관 동의 및 결제하기
    a = driver.find_element_by_xpath('//*[@id="cardPayment"]')
    driver.execute_script("arguments[0].click();", a)
    time.sleep(2)
    a1 = driver.find_element_by_xpath('//*[@id="container"]/section/article/section[1]/div/ul/li/span/label')
    driver.execute_script("arguments[0].click();", a1)
    time.sleep(2)
    a2 = driver.find_element_by_xpath('//*[@id="container"]/section/article/section[1]/div/a')
    driver.execute_script("arguments[0].click();", a2)
    time.sleep(2)

    # 포커스 이동
    tabs = driver.window_handles
    driver.switch_to.window(tabs[-1])
    driver.get_window_position(tabs[-1])

    #  약관 동의
    b = driver.find_element_by_xpath('//*[@id="prov_all"]')
    driver.execute_script("arguments[0].click();", b)
    time.sleep(2)

    # 카드 결제
    driver.find_element_by_xpath('//*[@id="P_CARD_NO"]').send_keys('9411170242904188')
    time.sleep(2)
    # driver.find_element_by_xpath('//*[@id="P_EXP_MON"]').send_keys(Keys.ENTER)
    # time.sleep(1)
    c = Select(driver.find_element_by_xpath('//*[@id="P_EXP_MON"]'))
    time.sleep(2)
    c.select_by_visible_text('11월')
    time.sleep(2)
    # driver.find_element_by_xpath('//*[@id="P_EXP_YEAR"]').send_keys(Keys.ENTER)
    # time.sleep(1)
    c1 = Select(driver.find_element_by_xpath('//*[@id="P_EXP_YEAR"]'))
    time.sleep(2)
    c1.select_by_visible_text('2020년')
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="P_CARD_PW"]').send_keys('05')
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="P_RR_NO"]').send_keys('931011')
    time.sleep(2)

    # 결제 완료
    b1 = driver.find_element_by_xpath('//*[@id="pay"]')
    driver.execute_script("arguments[0].click();", b1)
    time.sleep(2)

    # 서비스 이용페이지 이동
    driver.switch_to.window(tabs[1])
    driver.get_window_position(tabs[-1])
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="payment_end"]/div/div[2]/a').send_keys(Keys.ENTER)  # 서비스 이용 페이지로 이동
    time.sleep(2)


#############################################################결제 함수 실행##########################################################################
time.sleep(3)
pay()  # 결제 진입

##########################################################서비스 이용 함수 정의########################################################################
def service_check():
    driver.find_element_by_xpath('//*[@id="container"]/section[2]/article/section/div[1]/div[3]/a[2]').send_keys(Keys.ENTER)  # 보상 신청 페이지로 이동
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="siteId"]').send_keys('qoffice123')  # ID
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="siteOrderNo"]').send_keys('12345678')  # 주문번호
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="siteReturnFee"]').send_keys('5000')  # 반품비용
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="deliveryNo"]').send_keys('12345678')  # 운송장 번호
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="bankName"]').send_keys(Keys.ENTER)  # 은행명
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="container"]/section[2]/article/section[1]/div[4]/span[1]/span[1]/div/ul/li[6]/button').send_keys(Keys.ENTER)  # 농협은행 선택
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="bankAccountNo"]').send_keys('10004156035401')  # 계좌번호
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="bankUserName"]').send_keys('큐오피스')  # 예금주명
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="siteReason"]').send_keys('큐오피스 테스트')  # 반품사유
    time.sleep(1)

    agree_check2 = driver.find_element_by_xpath('//*[@id="chkboxAll"]')  # 약관 모두동의
    driver.execute_script("arguments[0].click();", agree_check2)
    time.sleep(2)

    reward_button = driver.find_element_by_xpath('//*[@id="submitBtn"]')  # 보상 신청 버튼 탭
    driver.execute_script("arguments[0].click();", reward_button)
    print("계좌 인증 발송")

    # 팝업 기다리기
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="check_account"]/div/a')))

    time.sleep(3)
    x_button = driver.find_element_by_xpath('//*[@id="check_account"]/div/ul/li[1]/a')  # 팝업 닫기
    driver.execute_script("arguments[0].click();", x_button)
    print("계좌 인증 성공")
    time.sleep(2)

##########################################################서비스 이용 함수 실행########################################################################
driver.get('https://interparkcare.co.kr/user/return_list')
time.sleep(2)
service_check()

##########################################################로그인/로그아웃 함수 실행#####################################################################
def loginout():
    # 로그인 페이지 진입
    driver.find_element_by_xpath('//*[@id="login_process"]').send_keys(Keys.ENTER)
    time.sleep(3)

    # 로그인 탭 이동
    driver.find_element_by_xpath('//*[@id="login"]/div/div[1]/button[2]').send_keys(Keys.ENTER)
    time.sleep(3)
    alert_ok()
    time.sleep(2)

    # 로그인 시도
    driver.find_element_by_id('email').send_keys('dmawl1592@gmail.com')  # 이메일 입력
    time.sleep(2)

    driver.find_element_by_id('password').send_keys('123456a!')  # 비밀번호 입력
    time.sleep(2)

    # 로그인 완료
    driver.find_element_by_xpath('//*[@id="login"]/div/div[2]/span[1]/button').send_keys(Keys.ENTER)
    time.sleep(2)
    alert_ok()
    time.sleep(2)
    print("로그인 성공")
    time.sleep(2)

    logout_button = driver.find_element_by_xpath('//*[@id="header"]/div/nav/div/ul/li[1]/a')  # 로그아웃
    driver.execute_script("arguments[0].click();", logout_button)
    alert_ok()
    print("로그아웃 성공")
    time.sleep(3)

###############################################################서비스 해지 실행########################################################################
def withdrawal():
    driver.get("https://interparkcare.co.kr/user/cancel")
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="container"]/section/article/section/ul/li[2]/a').send_keys(Keys.ENTER)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="cancel_end2"]/div/ul/li/a')))
    time.sleep(3)
    with_button = driver.find_element_by_xpath('//*[@id="cancel_end2"]/div/ul/li/a')
    driver.execute_script("arguments[0].click();", with_button)
    print("탈퇴 성공")

withdrawal()
time.sleep(2)
loginout()
time.sleep(3)

from selenium import webdriver
import chromedriver_autoinstaller as AutoChrome
from selenium.webdriver.common.by import By
import time
import pyperclip
import pyautogui
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#wait = WebDriverWait(driver, 10)
def start():
    global driver
    #크롬드라이버 버전 확인
    chrome_ver = AutoChrome.get_chrome_version().split('.')[0] 
    
    options = webdriver.ChromeOptions() # 브라우저 셋팅
    options.add_experimental_option("detach", True) # 브라우저 꺼짐 방지
    options.add_argument('lang=ko_KR') # 사용언어 한국어
    options.add_argument('disable-gpu') # 하드웨어 가속 안함
    options.add_experimental_option("excludeSwitches",['enable-logging']) # 불필요한 에러 메세지 삭제
    
    #실행 후 최신 버젼이 아니라서 실행이 안된다면 최신버젼으로 업데이트 후 재실행
    try:
        driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe', options = options)   
    except:
        AutoChrome.install(True)
        driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe', options = options)

    driver.implicitly_wait(10)

def naverLogin():
    # 네이버 로그인 주소 가져오기
    driver.get("https://nid.naver.com/nidlogin.login?mode=form&url=https%3A%2F%2Fwww.naver.com")
    # 화면 최대화
    driver.maximize_window()
    # F12 카피 셀렉터에서 태그 찾아오기
    # 아이디 입력
    id = driver.find_element(By.CSS_SELECTOR, '#id')
    id.click()
    pyperclip.copy('rhzn5512')
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    
    # 비밀번호 입력
    pw = driver.find_element(By.CSS_SELECTOR, '#pw')
    pw.click()
    pyperclip.copy('qhrtn5513!')
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    
    # 로그인 누르기
    driver.find_element(By.CSS_SELECTOR, '#log\.login').click()
    driver.implicitly_wait(10)

    # 등록 누르기
    driver.find_element(By.CSS_SELECTOR, '#new\.save').click()
    driver.implicitly_wait(10)
    
def mail(email, title, content):
    # 메일 들어가기 메일 보내기 누르기
    driver.get("https://mail.naver.com")
    driver.implicitly_wait(10)
    driver.find_element(By.CSS_SELECTOR, '#root > div > nav > div > div.lnb_header > div.lnb_task > a.item.button_write').click()
    driver.implicitly_wait(10)
    
    # 받는 사람
    driver.find_element(By.CSS_SELECTOR, '#user_input_1').send_keys(email)
    time.sleep(1)
    
    # 제목 입력
    driver.find_element(By.CSS_SELECTOR, '#subject_title').send_keys(title)
    time.sleep(1)
    
    # 본문 작성 (iframe 안으로 들어가기)
    iframe = driver.find_element(By.CSS_SELECTOR, '#content > div.contents_area > div > div.editor_area > div > div.editor_body > iframe')
    driver.switch_to.frame(iframe)
    body = driver.find_element(By.CSS_SELECTOR,'body > div > div.workseditor-content')
    body.send_keys(content)
    time.sleep(1)
    
    # frame 밖으로 나가기
    driver.switch_to.default_content()
    
    # 메일 보내기
    driver.find_element(By.CSS_SELECTOR, '#content > div.mail_toolbar.type_write > div:nth-child(1) > div > button.button_write_task').click()
    driver.implicitly_wait(10)
    
if __name__ == '__main__':
    start()
    naverLogin()
    mail('rhzn5512@naver.com', '자동화 타이틀', '테스트용 본문 내용입니다.')
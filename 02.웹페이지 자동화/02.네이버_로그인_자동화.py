from selenium import webdriver
import chromedriver_autoinstaller as AutoChrome
from selenium.webdriver.common.by import By
import time
import pyperclip
import pyautogui

def start():
    global driver
    #크롬드라이버 버전 확인
    chrome_ver = AutoChrome.get_chrome_version().split('.')[0] 
    #브라우저 꺼짐 방지
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    #불필요한 에러 메세지 삭제
    options.add_experimental_option("excludeSwitches",['enable-logging'])
    
    #실행 후 최신 버젼이 아니라서 실행이 안된다면 최신버젼으로 업데이트 후 재실행
    try:
        driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe', options = options)   
    except:
        AutoChrome.install(True)
        driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe', options = options)

    driver.implicitly_wait(10)

if __name__ == '__main__':
    start()
    # 주소 가져오기
    driver.get("https://nid.naver.com/nidlogin.login?mode=form&url=https%3A%2F%2Fwww.naver.com")
    # 화면 최대화
    driver.maximize_window()
    # F12 카피 셀렉터에서 태그 찾아오기
    # 아이디 입력창
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
    
    driver.find_element(By.CSS_SELECTOR, '#log\.login').click()
    time.sleep(1)
    
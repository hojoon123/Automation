from selenium import webdriver
import chromedriver_autoinstaller as AutoChrome

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
    driver.get("http://www.naver.com")
    
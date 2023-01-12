import pyautogui
import time
import pyperclip

#1. 키보드 입력(문자)
pyautogui.write('startcoding', interval = 0.1)

#2. 키보드 입력(키)
pyautogui.press('enter')

#3. 키보드 입력(동시입력)
pyautogui.hotkey('ctrl', 'a')

#4. 한글 입력 방법
pyperclip.copy('한글입력')
pyautogui.hotkey('ctrl', 'v')

#5. 문자열 출력하기
print(pyperclip.paste())

#6. 키보드 제어 중단
#ctrl + alt + del 윈도우 OS 전용

#7. 현재 화면의 스크린 사이즈 가져오기
size = pyautogui.size()
print(size)
# size[0] == width
# size[1] == height


# 스크린샷 찍기
img = pyautogui.screenshot()

# 파일 저장
img.save("screenshot.png")


#화면 특정좌표의 픽셀값 가져오기

pixel = pyautogui.pixel(0, 0)
# RGB값 출력 ex) (60, 64, 67)
print(pixel)


#화면 특정좌표의 픽셀값 색상 일치하는지 검사
# if pyautogui.pixelMatchesColor(x좌표, y좌표, (R, G, B)):
if pyautogui.pixelMatchesColor(0, 0, (60, 64, 67)):
    print("색상 일치")
else:
    print("색상 불일치")



#화면에 이미지 존재하는지 확인 (locateOnScreen)

target_img= pyautogui.locateOnScreen("target_img.png")
print(target_img)
# 이미지를 찾으면 결과는 다음과 같이 나온다. ex) Box(left=900, top=82, width=31, height=23)
# 이미지를 찾지 못하면 결과는 None 이라고 나온다

# 이미지 클릭
pyautogui.click(target_img)

# 이미지 찾아서 마우스 이동
pyautogui.moveTo(target_img)



# 1. GrayScale 로 찾기
# 흑백으로 전환해서 찾기. 정확도가 떨어지지만 30% 정도 속도가 개선됨
target_img= pyautogui.locateOnScreen("target_img.png", grayscale=True)
print(target_img)



# 2. 범위 지정해서 찾기
# 메뉴같은 경우 위쪽에만 있는 상황에서 사용
# target_img= pyautogui.locateOnScreen("target_img.png", region=(x, y, width, height))
target_img= pyautogui.locateOnScreen("target_img.png", region=(0, 0, 1024, 300))
print(target_img)



# 3. 정확도를 낮춰서 찾기
# 분명히 똑같은 이미지로 자동화했는데 사소한 차이로 클릭이 안되는 상황에서 사용
# 주의) pip install opencv-python 명령어로 패키지 설치해야 한다.

# confidence=0.999 값이 기본값이다(=99.9%).
# 아래 코드에서는 80% 정확도로 해본다.
target_img= pyautogui.locateOnScreen("target_img.png", confidence=0.8)
print(target_img)


#https://blog.naver.com/PostView.naver?blogId=bb_&logNo=222630255538&parentCategoryNo=&categoryNo=103&viewDate=&isShowPopularPosts=true&from=search
#착한흑곰 블로그 여기서 걍 찾아보셈
'''
문자열

의미

'a', 'b', 'c', 'A', 'B', 'C', '1', '2', '3', '!', '@', '#' 등등

해당하는 단일문자

'enter' (또는 'return' 또는 '\n')

ENTER 키

'esc'

ESC 키

'shiftleft', 'shiftright'

왼쪽과 오른쪽 SHIFT 키

'altleft', 'altright'

왼쪽과 오른쪽 ALT 키

'ctrlleft', 'ctrlright'

왼쪽과 오른쪽 CTRL 키

'tab' (또는 '\t')

TAB 키

'backspace', 'delete'

BACKSPACE 키, DELETE 키

'pageup', 'pagedown'

PAGE UP 키, PAGE DOWN 키

'home', 'end'

HOME 키, END 키

'up', 'down', 'left', 'right'

상, 하, 좌, 우 화살표 키

'f1', 'f2', 'f3' 등등

F1 ~ F12 키

'volumemute', 'volumedown', 'volumeup'

음소거, 볼륨 감소 키, 볼륨 증가 키

(일부 키보드에는 이러한 키가 없지만 운영 체제에서는 시뮬레이션된 키입력을 이해할 수 있음)

'pause'

PAUSE 키

'capslock', 'numlock', 'scrolllock'

CAPS LOCK 키, NUM LOCK 키, SCROLL LOCK 키

'insert'

INS 또는 INSERT 키

'printscreen'

PRTSC 또는 PRINT SCREEN 키

'winleft', 'winright'

왼쪽과 오른쪽 윈도우(WIN) 키 (Windows OS 전용)
'''

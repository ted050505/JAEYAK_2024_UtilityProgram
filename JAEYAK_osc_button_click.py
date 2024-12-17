from pythonosc import dispatcher
from pythonosc import osc_server
import pyautogui
import win32gui

# 버튼 클릭 함수
def click_with_pyautogui(window_title, relative_x, relative_y):
    hwnd = win32gui.FindWindow(None, window_title)
    if hwnd == 0:
        print(f"창 '{window_title}'을(를) 찾을 수 없습니다.")
        return

    # 창의 절대 좌표 가져오기
    rect = win32gui.GetWindowRect(hwnd)
    absolute_x = rect[0] + relative_x
    absolute_y = rect[1] + relative_y
    print(f"Calculated Absolute Position: X={absolute_x}, Y={absolute_y}")

    # PyAutoGUI를 사용한 클릭
    pyautogui.moveTo(absolute_x, absolute_y)
    pyautogui.click()
    print("Button clicked successfully!")

# OSC 메시지를 처리하는 함수
def on_osc_message(address, *args):
    print(f"Received: {address}, {args}")
    if address == "/trigger" and args[0] == "button_click":
        click_with_pyautogui("Dollars_DEEP_Lite", 100, 44)

# 프로그램 초기 설정
def setup_program():
    global port, window_title

    while True:
        try:
            port = int(input("Enter the port number to use (default: 8000): ") or 8000)
            break
        except ValueError:
            print("Invalid input. Please enter a valid port number.")

    # 제어할 창 이름 입력
    window_title = input("Enter the title of the window to control (default: 'Dollars_DEEP_Lite'): ") or "Dollars_DEEP_Lite"
    print(f"Target window set to: '{window_title}'")

# 프로그램 초기화
setup_program()

# OSC 서버 설정
dispatcher = dispatcher.Dispatcher()
dispatcher.map("/trigger", on_osc_message)

ip = "127.0.0.1"
port = 8000
server = osc_server.ThreadingOSCUDPServer((ip, port), dispatcher)

print(f"OSC server listening on {ip}:{port}")
server.serve_forever()

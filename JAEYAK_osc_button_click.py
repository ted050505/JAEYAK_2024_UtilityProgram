from pythonosc import dispatcher
from pythonosc import osc_server
import pyautogui
import win32gui
import json
import os
import time

class WindowController:
    def __init__(self):
        self.config_file = 'window_config.json'
        self.setup_program()
        self.setup_button_position()

    def setup_program(self):
        while True:
            try:
                port_input = input("사용할 포트 번호를 입력하세요 (기본값: 8000): ") or "8000"
                self.port = int(port_input)
                break
            except ValueError:
                print("올바르지 않은 입력입니다. 유효한 포트 번호를 입력해주세요.")

        self.window_title = input("제어할 창의 이름을 입력하세요 (기본값: 'Dollars_DEEP_Lite'): ") or "Dollars_DEEP_Lite"
        print(f"대상 창이 '{self.window_title}'(으)로 설정되었습니다.")

        self.save_config()

    def setup_button_position(self):
        print("\n버튼 위치 설정을 시작합니다...")
        print("대상 창을 활성화하고 버튼이 보이도록 준비해주세요.")
        input("준비가 되면 Enter를 눌러주세요...")

        hwnd = win32gui.FindWindow(None, self.window_title)
        if hwnd == 0:
            print(f"창 '{self.window_title}'을(를) 찾을 수 없습니다.")
            return

        # 창이 최소화되어 있다면 복원
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE

        # 창의 크기와 위치 정보 가져오기
        window_rect = win32gui.GetWindowRect(hwnd)
        window_x = window_rect[0]
        window_y = window_rect[1]
        window_width = window_rect[2] - window_rect[0]
        window_height = window_rect[3] - window_rect[1]

        print("\n버튼 위치 설정 방법:")
        print("1. 마우스를 클릭하려는 버튼 위에 올려놓으세요.")
        print("2. 3초 후에 마우스 커서의 위치를 자동으로 저장합니다.")
        print("3초 카운트다운을 시작합니다...")
        
        for i in range(3, 0, -1):
            print(f"{i}...")
            time.sleep(1)

        # 현재 마우스 커서 위치 가져오기
        cursor_x, cursor_y = pyautogui.position()

        # 상대 위치 계산
        relative_x = cursor_x - window_x
        relative_y = cursor_y - window_y

        # 위치 정보를 비율로 저장
        self.button_x_ratio = relative_x / window_width
        self.button_y_ratio = relative_y / window_height

        print(f"\n버튼 위치가 저장되었습니다:")
        print(f"창 크기 - 너비: {window_width}, 높이: {window_height}")
        print(f"상대 위치 - X: {relative_x}, Y: {relative_y}")
        print(f"비율 - X: {self.button_x_ratio:.3f}, Y: {self.button_y_ratio:.3f}")

        # 설정 저장
        self.save_config()

    def save_config(self):
        config = {
            'port': self.port,
            'window_title': self.window_title,
            'button_x_ratio': getattr(self, 'button_x_ratio', 0.5),
            'button_y_ratio': getattr(self, 'button_y_ratio', 0.5)
        }
        try:
            with open(self.config_file, 'w') as f:
                json.dump(config, f)
        except Exception as e:
            print(f"설정 저장 중 오류 발생: {e}")

    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                self.port = config.get('port', 8000)
                self.window_title = config.get('window_title', 'Dollars_DEEP_Lite')
                self.button_x_ratio = config.get('button_x_ratio', 0.5)
                self.button_y_ratio = config.get('button_y_ratio', 0.5)
            except Exception as e:
                print(f"설정 파일 로드 중 오류 발생: {e}")
                self.set_default_values()
        else:
            self.set_default_values()

    def set_default_values(self):
        self.port = 8000
        self.window_title = 'Dollars_DEEP_Lite'

    def click_with_pyautogui(self):
        try:
            hwnd = win32gui.FindWindow(None, self.window_title)
            if hwnd == 0:
                print(f"창 '{self.window_title}'을(를) 찾을 수 없습니다.")
                return False

            # 창이 최소화되어 있는지 확인
            if win32gui.IsIconic(hwnd):
                print("창이 최소화되어 있습니다. 복원합니다.")
                win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE

            # 현재 창 크기 가져오기
            rect = win32gui.GetWindowRect(hwnd)
            window_width = rect[2] - rect[0]
            window_height = rect[3] - rect[1]

            # 비율을 기반으로 실제 클릭 위치 계산
            relative_x = int(window_width * self.button_x_ratio)
            relative_y = int(window_height * self.button_y_ratio)
            absolute_x = rect[0] + relative_x
            absolute_y = rect[1] + relative_y

            # 화면 범위 검증
            screen_width, screen_height = pyautogui.size()
            if not (0 <= absolute_x <= screen_width and 0 <= absolute_y <= screen_height):
                print("계산된 좌표가 화면 범위를 벗어났습니다.")
                return False

            print(f"클릭 위치: X={absolute_x}, Y={absolute_y}")
            pyautogui.moveTo(absolute_x, absolute_y)
            pyautogui.click()
            return True

        except Exception as e:
            print(f"클릭 작업 중 오류 발생: {e}")
            return False

# OSC 메시지 핸들러
def on_osc_message(address, *args):
    print(f"수신: {address}, {args}")
    if address == "/trigger" and args[0] == "button_click":
        controller.click_with_pyautogui()

# 메인 실행 부분
if __name__ == "__main__":
    controller = WindowController()
    
    dispatcher = dispatcher.Dispatcher()
    dispatcher.map("/trigger", on_osc_message)

    try:
        server = osc_server.ThreadingOSCUDPServer(
            ("127.0.0.1", controller.port), dispatcher)
        print(f"OSC 서버 시작됨 - {server.server_address[0]}:{server.server_address[1]}")
        server.serve_forever()
    except Exception as e:
        print(f"서버 시작 중 오류 발생: {e}")

from pywinauto import Application
import pyautogui

# 프로그램 창 연결 (창 제목 입력)
app = Application(backend="uia").connect(title="Dollars_DEEP_Lite")
window = app.window(title="Dollars_DEEP_Lite")

# 창 활성화
window.set_focus()
print(window.is_active())

# # 2. 창의 위치와 크기 확인 (디버깅 용도)
# rect = window.rectangle()
# print(f"Window Position: Left={rect.left}, Top={rect.top}, Right={rect.right}, Bottom={rect.bottom}")

# 1. 창의 절대 좌표 가져오기
rect = window.rectangle()
print(f"Window Position: Left={rect.left}, Top={rect.top}, Right={rect.right}, Bottom={rect.bottom}")

# 2. 마우스 위치로 버튼 절대 좌표 확인
print("마우스를 버튼 위에 올려놓고 현재 위치를 확인하세요.")
input("Enter 키를 눌러서 확인: ")  # 잠시 멈춰 사용자가 마우스 위치를 조정할 시간 제공
mouse_x, mouse_y = pyautogui.position()
print(f"Mouse Absolute Position: X={mouse_x}, Y={mouse_y}")

# 3. 버튼의 상대 좌표 계산
button_relative_x = mouse_x - rect.left
button_relative_y = mouse_y - rect.top
print(f"Button Relative Position: X={button_relative_x}, Y={button_relative_y}")

# 버튼의 절대 좌표 계산 (창의 위치 + 상대 좌표)
button_absolute_x = rect.left + button_relative_x
button_absolute_y = rect.top + button_relative_y

# 창 활성화 및 버튼 클릭
window.set_focus()
pyautogui.click(x=button_absolute_x, y=button_absolute_y)
print("Button clicked successfully!")

# 내부 컨트롤 식별 출력
# window.print_control_identifiers()

# 내부 컨트롤 계층 구조 출력
# window.dump_tree()
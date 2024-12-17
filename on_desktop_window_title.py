from pywinauto import Desktop

# 현재 활성 창 목록 출력
windows = Desktop(backend="uia").windows()
for win in windows:
    print(f"Title: {win.window_text()} | Handle: {win.handle}")

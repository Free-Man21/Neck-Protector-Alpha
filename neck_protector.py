from tkinter import *
from tkinter import messagebox
from win10toast import ToastNotifier
from PIL import Image
import os
import sys
import threading
import time
import pystray
import winreg
import win32event
import win32api
import pywintypes
import locale

background_thread = None
stop_event = threading.Event()
current_time = None
settings_win = None

def get_system_language():
    lang = locale.getdefaultlocale()[0]
    if lang is not None and lang.startswith("ko"):
        return "ko"
    else:
        return "en"
    
LANG = get_system_language()

STRINGS = {
    "ko": {
        "tutorial_title": "Neck Protector 사용법",
        "tutorial_text": (
            "이 프로그램은 일정 시간마다 알림을 보내 목 스트레칭을 도와줍니다.\n\n"
            "※ 주의: 특정 앱(게임, 전체화면 등)이 실행 중일 때는 알림이 표시되지 않을 수 있습니다.\n"
            "이 경우, 윈도우 알림 우선순위 설정에서 'Neck Protector'를 허용 앱으로 추가해 주세요.\n\n"
            "아래 버튼을 눌러 알림 우선순위 설정 화면으로 이동할 수 있습니다."
        ),
        "open_priority": "알림 우선순위 설정 열기",
        "open": "열기",
        "close": "닫기",
        "exit": "종료",
        "settings": "설정",
        "usage": "사용법",
        "save_confirm": "이 상태로 저장할까요?",
        "save_confirm_title": "저장 확인",
        "save": "저장",
        "cancel": "취소",
        "main_title": "Neck Protector"
    },
    "en": {
        "tutorial_title": "How to use Neck Protector",
        "tutorial_text": (
            "This program sends notifications to help you stretch your neck at regular intervals.\n\n"
            "※ Note: Notifications may not appear while certain apps (games, fullscreen, etc.) are running.\n"
            "In this case, please add 'Neck Protector' as an allowed app in Windows notification priority settings.\n\n"
            "Click the button below to open the notification priority settings screen."
        ),
        "open_priority": "Open Priority Notification Settings",
        "open": "Open",
        "close": "Close",
        "exit": "Exit",
        "settings": "Settings",
        "usage": "Usage",
        "save_confirm": "Save changes?",
        "save_confirm_title": "Save Confirmation",
        "save": "Save",
        "cancel": "Cancel",
        "main_title": "Neck Protector"
    }
}

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def get_app_data_path():
    if sys.platform == "win32":
        return os.path.join(os.environ.get("LOCALAPPDATA"), "Neck Protector")
    else:
        return os.path.abspath(".")

def convert_png_to_ico_and_get_path():
    app_data_dir = get_app_data_path()
    os.makedirs(app_data_dir, exist_ok=True)

    png_path = resource_path("neck_protector.png")
    ico_output_path = os.path.join(app_data_dir, "neck_protector_toast_icon.ico")

    try:
        img = Image.open(png_path)
        img.save(ico_output_path, sizes=[(16,16), (24,24), (32,32), (48,48), (64,64)])
        return ico_output_path
    except Exception as e:
        print(f"Failed to convert PNG to ICO for toast: {e}")
        return "" # 변환 실패 시 아이콘 없이 진행

def get_saved_time_path():
    app_data_dir = get_app_data_path()
    os.makedirs(app_data_dir, exist_ok=True)
    return os.path.join(app_data_dir, "saved_time.txt")

def add_to_startup_registry():
    exe_path = f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}" --background'
    key = r"Software\Microsoft\Windows\CurrentVersion\Run"
    app_name = "NeckProtector"
    try:
        reg_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_SET_VALUE)
    except FileNotFoundError:
        reg_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key)
    winreg.SetValueEx(reg_key, app_name, 0, winreg.REG_SZ, exe_path)
    winreg.CloseKey(reg_key)

def background_task(selected_time, stop_event):
    toaster = ToastNotifier()
    while not stop_event.is_set():
        for _ in range(selected_time * 60):
            if stop_event.is_set():
                return
            time.sleep(1)
        toaster.show_toast(
            "Neck Protector",
            "목 스트레칭 시간입니다!",
            duration=10,
            threaded=True,
            icon_path=toaster_icon_path # 이 줄로 다시 변경합니다.
        )

def start_background_task(selected_time):
    global background_thread, stop_event
    # 기존 쓰레드 종료
    if background_thread and background_thread.is_alive():
        stop_event.set()
        background_thread.join()
    stop_event.clear()
    background_thread = threading.Thread(target=background_task, args=(selected_time, stop_event), daemon=True)
    background_thread.start()

def on_closing():
    window.withdraw()  # 창 숨기기

def show_tray_icon():
    def on_exit(icon, item):
        stop_event.set()
        icon.stop()
        window.quit()  # 완전히 종료

    def on_settings(icon, item):
        window.after(0, lambda: [window.deiconify(), window.lift(), window.focus_force()])

    image = Image.open(resource_path("neck_protector.png")).resize((64, 64))
    menu = pystray.Menu(
        pystray.MenuItem(STRINGS[LANG]["open"], on_settings),
        pystray.MenuItem(STRINGS[LANG]["exit"], on_exit)
    )
    icon = pystray.Icon("neck_protector", image, STRINGS[LANG]["main_title"], menu)
    threading.Thread(target=icon.run_detached, daemon=True).start()

def get_saved_time():
    try:
        with open(get_saved_time_path(), "r") as f:
            return int(f.read())
    except:
        return 2  # 기본값
    
def open_settings_window():
    global settings_win
    if settings_win is not None and settings_win.winfo_exists():
        settings_win.lift()
        settings_win.focus_force()
        return

    def save_and_close():
        selected_time = settings_time_slider.get()
        with open(get_saved_time_path(), "w") as f:
            f.write(str(selected_time))
        global current_time, settings_win
        if selected_time != current_time:
            current_time = selected_time
            start_background_task(current_time)
        settings_win.destroy()
        settings_win = None  # 창 닫힐 때 변수 초기화

    def close_without_save():
        global settings_win
        settings_win.destroy()
        settings_win = None

    def on_close():
        # 경고창 띄우기
        result = messagebox.askyesno(
            STRINGS[LANG]["save_confirm_title"],
            STRINGS[LANG]["save_confirm"]
        )
        if result:  # 예
            save_and_close()
        else:       # 아니요
            close_without_save()

    settings_win = Toplevel(window)
    settings_win.title(STRINGS[LANG]["settings"])
    win_width, win_height = 400, 300
    screen_width = settings_win.winfo_screenwidth()
    screen_height = settings_win.winfo_screenheight()
    x = (screen_width // 2) - (win_width // 2)
    y = (screen_height // 2) - (win_height // 2)
    settings_win.geometry(f"{win_width}x{win_height}+{x}+{y}")
    settings_win.resizable(False, False)
    settings_win.configure(bg="black")
    settings_win.protocol("WM_DELETE_WINDOW", on_close)
    settings_win.transient(window)

    Label(settings_win, text="시간(분) 설정", font=("Arial", 14), bg="black", fg="white").pack(pady=(30, 10))
    settings_time_slider = Scale(
        settings_win,
        from_=2,
        to=30,
        orient=HORIZONTAL,
        length=280,
        font=("Arial", 12),
        bg="black",
        fg="white",
        troughcolor="lightblue",
        highlightthickness=0
    )
    settings_time_slider.set(current_time)
    settings_time_slider.pack(pady=10)

    btn_frame = Frame(settings_win, bg="black")
    btn_frame.pack(pady=30)

    Button(
        btn_frame, text=STRINGS[LANG]["save"], command=save_and_close,
        bg="black", fg="white", font=("Arial", 12, "bold"),
        activebackground="black", activeforeground="white",
        relief=FLAT, padx=20, pady=8
    ).pack(side=LEFT, padx=10)

    Button(
        btn_frame, text=STRINGS[LANG]["cancel"], command=close_without_save,
        bg="black", fg="white", font=("Arial", 12, "bold"),
        activebackground="black", activeforeground="white",
        relief=FLAT, padx=16, pady=8
    ).pack(side=LEFT, padx=10)

def show_tutorial_window():
    tutorial_win = Toplevel(window)
    tutorial_win.title(STRINGS[LANG]["tutorial_title"])
    win_width, win_height = 420, 320
    screen_width = tutorial_win.winfo_screenwidth()
    screen_height = tutorial_win.winfo_screenheight()
    x = (screen_width // 2) - (win_width // 2)
    y = (screen_height // 2) - (win_height // 2)
    tutorial_win.geometry(f"{win_width}x{win_height}+{x}+{y}")
    tutorial_win.resizable(False, False)
    tutorial_win.configure(bg="black")
    tutorial_win.transient(window)
    tutorial_win.grab_set()

    Label(
        tutorial_win, text=STRINGS[LANG]["tutorial_text"],
        font=("Arial", 12), bg="black", fg="white", justify="left", wraplength=370
    ).pack(padx=24, pady=24)

    def open_windows_notification_settings():
        import subprocess
        subprocess.Popen("start ms-settings:quiethours-priorities-apps", shell=True)

    Button(
        tutorial_win, text=STRINGS[LANG]["open_priority"], command=open_windows_notification_settings,
        bg="#1976d2", fg="white", font=("Arial", 12, "bold"),
        activebackground="#1565c0", activeforeground="white",
        relief=FLAT, padx=20, pady=8
    ).pack(pady=(0, 10))


window = Tk()
window.title(STRINGS[LANG]["main_title"])
app_icon_image = PhotoImage(file=resource_path("neck_protector.png"))
window.iconphoto(True, app_icon_image)

# 중앙 정렬 코드 추가
win_width, win_height = 600, 600
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()
x = (screen_width // 2) - (win_width // 2)
y = (screen_height // 2) - (win_height // 2)
window.geometry(f"{win_width}x{win_height}+{x}+{y}")

window.resizable(False, False)
window.configure(bg="black")
window.protocol("WM_DELETE_WINDOW", on_closing)
window.attributes('-toolwindow', True)

# UI 구성
main_frame = Frame(window, bg="black")
main_frame.pack(expand=True)

title_label = Label(main_frame, text=STRINGS[LANG]["main_title"], font=("Arial", 22, "bold"), bg="black", fg="white")
title_label.pack(pady=(30, 10), side=TOP)

icon_img = PhotoImage(file=resource_path("neck_protector.png"))
icon_img = icon_img.subsample(3, 3)
icon_label = Label(main_frame, image=icon_img, bg="lightblue")
icon_label.pack(pady=(30, 30), side=TOP)

settings_btn = Button(
    main_frame,
    text=STRINGS[LANG]["settings"],
    command=open_settings_window,
    bg="black",
    fg="white",
    font=("Arial", 12, "bold"),
    activebackground="black",
    activeforeground="white",
    relief=FLAT,
    padx=16,
    pady=8
)
settings_btn.pack(pady=10, side=BOTTOM)

tutorial_btn = Button(
    main_frame,
    text=STRINGS[LANG]["usage"],
    command=show_tutorial_window,
    bg="black",
    fg="white",
    font=("Arial", 12, "bold"),
    activebackground="black",
    activeforeground="white",
    relief=FLAT,
    padx=16,
    pady=8
)
tutorial_btn.pack(pady=10, side=BOTTOM)

mutex_name = "NeckProtectorSingleInstanceMutex"
try:
    handle = win32event.CreateMutex(None, False, mutex_name)
    # 이미 실행 중이면 WAIT_ABANDONED(0x00000080) 또는 WAIT_OBJECT_0(0) 이외의 값이 반환됨
    if win32api.GetLastError() == 183:
        # 이미 실행 중인 인스턴스가 있음
        # 창을 띄우는 대신 바로 종료
        sys.exit(0)
except pywintypes.error:
    # 예외 발생 시 그냥 종료
    sys.exit(0)

if __name__ == "__main__":
    global toaster_icon_path
    toaster_icon_path = convert_png_to_ico_and_get_path()

    add_to_startup_registry()

    current_time = get_saved_time()
    show_tray_icon()
    start_background_task(current_time)

    if "--background" in sys.argv:
        window.withdraw()
    else:
        window.deiconify()

    window.mainloop()
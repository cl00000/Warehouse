# function/function_autoPrint.py
import time
import pyautogui
import win32gui
import win32con
from PIL import ImageGrab
import threading
from ctypes import windll
from typing import Optional, Callable
import winsound

# 配置参数
monitor_region = (1018, 254, 160, 23)  # 检测区域 (x, y, width, height)
click_sequence = [
    (1000, 776, 0.5),  # 第一次点击: (x,y), 等待1秒后执行
    (949, 600, 0.1),  # 第二次点击: (x,y), 上次点击后等待0.1秒
    (732, 155, 0.3)  # 第三次点击: (x,y), 上次点击后等待0.3秒
]
check_interval = 0.4  # 检测频率(秒)

# 线程控制
monitor_thread = None
stop_event = threading.Event()
lock = threading.Lock()

# 错误回调函数
error_callback: Optional[Callable[[str], None]] = None

def set_error_callback(callback: Callable[[str], None]):
    """设置错误回调函数"""
    global error_callback
    error_callback = callback

def _notify_error(message: str):
    """通知错误信息"""
    if error_callback:
        error_callback(message)

def find_child_window(parent_title: str, child_title: str) -> Optional[int]:
    """查找指定父窗口下的子窗口"""
    try:
        def callback(hwnd: int, hwnds: list):
            if win32gui.IsWindowVisible(hwnd) and child_title.lower() in win32gui.GetWindowText(hwnd).lower():
                hwnds.append(hwnd)
            return True

        parent_hwnd = win32gui.FindWindow(None, parent_title)
        if not parent_hwnd:
            _notify_error(f"未找到父窗口: {parent_title}")
            return None

        child_hwnds = []
        win32gui.EnumChildWindows(parent_hwnd, callback, child_hwnds)
        return child_hwnds[0] if child_hwnds else None
    except Exception as e:
        _notify_error(f"查找窗口出错: {str(e)}")
        return None

def is_window_active() -> bool:
    """检查目标窗口是否存在且激活"""
    hwnd = find_child_window("旺店通ERP", "发货确认")
    if not hwnd:
        _notify_error("未找到'旺店通ERP'的'发货确认'窗口")
    return bool(hwnd)

def bring_window_to_front() -> bool:
    """将目标窗口置顶，返回是否成功"""
    hwnd = find_child_window("旺店通ERP", "发货确认")
    if not hwnd:
        return False

    try:
        # 先尝试正常置顶方式
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)

        # 特殊处理被其他窗口遮挡的情况
        if windll.user32.GetForegroundWindow() != hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            time.sleep(0.1)
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        return True
    except Exception as e:
        _notify_error(f"窗口置顶失败: {str(e)}")
        return False

def monitor_task():
    """监控任务主循环"""
    try:
        previous_image = ImageGrab.grab(bbox=(
            monitor_region[0],
            monitor_region[1],
            monitor_region[0] + monitor_region[2],
            monitor_region[1] + monitor_region[3]
        ))

        while not stop_event.is_set():
            if not is_window_active():
                _notify_error("目标窗口未激活，暂停监控...")
                time.sleep(1)
                continue

            current_image = ImageGrab.grab(bbox=(
                monitor_region[0],
                monitor_region[1],
                monitor_region[0] + monitor_region[2],
                monitor_region[1] + monitor_region[3]
            ))

            if list(current_image.getdata()) != list(previous_image.getdata()):
                print("检测到区域变化，执行点击序列...")

                if not bring_window_to_front():  # 确保窗口在前
                    continue

                for i, (x, y, delay) in enumerate(click_sequence, 1):
                    if stop_event.is_set():
                        break
                    time.sleep(delay)
                    pyautogui.click(x, y)
                    print(f"第{i}次点击: 位置 ({x},{y}) 延迟 {delay}秒")

                # 在最后一次点击后播放提示音
                print("最后一次点击，播放成功提示音...")
                winsound.Beep(1000, 500)  # 播放1000Hz的音调，持续500ms

            previous_image = current_image
            time.sleep(check_interval)

    except Exception as e:
        _notify_error(f"监控任务异常: {str(e)}")
    finally:
        print("监控任务已退出")
        _notify_error("监控任务已停止")

def start_auto_print() -> bool:
    """启动监控线程，返回是否成功"""
    global monitor_thread
    with lock:
        if monitor_thread and monitor_thread.is_alive():
            _notify_error("监控已在运行中")
            return False

        if not is_window_active():
            _notify_error("错误：未找到'旺店通ERP'的'发货确认'窗口")
            return False

        stop_event.clear()
        monitor_thread = threading.Thread(target=monitor_task, daemon=True)
        monitor_thread.start()
        print("监控已启动")
        return True

def stop_auto_print():
    """停止监控线程"""
    global monitor_thread
    with lock:
        if monitor_thread and monitor_thread.is_alive():
            stop_event.set()
            monitor_thread.join(timeout=1)
            print("监控已停止")
        monitor_thread = None
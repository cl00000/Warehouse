# function/function_switch2.py
import time
import threading
import pyautogui
import winsound
from function_counter import counter_manager
from function_color_detector import get_color_at_region, detect_color

CLICK_POSITION = (650, 983)  # 点击位置：退货按钮

class AutoWarehouseMonitor:
    """自动入库监控器"""

    POLL_INTERVAL = 0.1  # 检测间隔(秒)

    def __init__(self):
        self._stop_event = threading.Event()
        self._monitor_thread = None

    def start(self):
        """启动监控"""
        if self._monitor_thread and self._monitor_thread.is_alive():
            return False

        self._stop_event.clear()
        self._monitor_thread = threading.Thread(
            target=self._monitor_loop,
            daemon=True
        )
        self._monitor_thread.start()
        return True

    def stop(self):
        """停止监控"""
        self._stop_event.set()
        if self._monitor_thread:
            self._monitor_thread.join(timeout=1)

    def _monitor_loop(self):
        """监控循环"""
        last_color = None

        while not self._stop_event.is_set():
            current_color = get_color_at_region()
            detected_color = detect_color(current_color)

            if detected_color != last_color:
                last_color = detected_color
                if detected_color == "green":
                    play_sound()
                    pyautogui.click(*CLICK_POSITION)
                    counter_manager.increment()

            time.sleep(self.POLL_INTERVAL)

    def is_running(self):
        """检查监控状态"""
        return (self._monitor_thread and
                self._monitor_thread.is_alive() and
                not self._stop_event.is_set())


# 全局实例
_monitor = AutoWarehouseMonitor()

def play_sound():
    """播放提示音"""
    winsound.Beep(2500 ,200)

def start_monitoring():
    """启动监控"""
    print('自动入库已启动')
    return _monitor.start()


def stop_monitoring():
    # print('自动入库已停止')
    """停止监控"""
    _monitor.stop()


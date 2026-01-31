# function/function_keyboard_manager.py
import threading
import time

import keyboard
import pyautogui
from typing import Optional
from PySide6.QtCore import QObject, Signal
from function_counter import counter_manager


class KeyboardManager(QObject):
    """
    键盘映射管理器

    功能：
    1. 空格键映射到点击(650, 983)坐标
    2. Alt键触发点击清空按钮
    """
    # 新增信号
    left_key_pressed = Signal()
    right_key_pressed = Signal()

    def __init__(self):
        super().__init__()  # 需要调用父类初始化
        self._space_remapped: bool = False
        self._delete_listener: Optional[keyboard.HotKey] = None
        self._click_position: tuple[int, int] = (123, 189)  # 可配置的点击坐标
        self._space_click_position: tuple[int, int] = (650, 983)  # 空格键点击坐标

        self._left_listener: Optional[keyboard.HotKey] = None
        self._right_listener: Optional[keyboard.HotKey] = None

    def enable_keyboard_mapping(self):
        """启用所有键盘映射功能"""
        self.enable_space_click()
        self.enable_alt_click()
        self._enable_arrow_keys()

    def disable_keyboard_mapping(self):
        """禁用所有键盘映射功能"""
        self.disable_space_click()
        self.disable_alt_click()
        self._disable_arrow_keys()

    def _enable_arrow_keys(self):
        """启用方向键监听（仅主键盘区）"""
        try:
            # 使用hook替代on_press_key以获取更多事件信息
            self._hook = keyboard.hook(self._on_key_event)
        except Exception as e:
            print(f"启用方向键监听失败: {e}")

    def _disable_arrow_keys(self):
        """禁用方向键监听"""
        if hasattr(self, "_hook") and self._hook is not None:
            keyboard.unhook(self._hook)
            self._hook = None

    def _on_key_event(self, event):
        """处理按键事件，仅响应主键盘方向键"""
        if event.event_type == keyboard.KEY_DOWN:
            # Windows 下主键盘左箭头扫描码为 75，右箭头为 77
            if event.name == 'left' and event.scan_code == 75:
                self.left_key_pressed.emit()
                return False  # 抑制主键盘左箭头事件
            elif event.name == 'right' and event.scan_code == 77:
                self.right_key_pressed.emit()
                return False  # 抑制主键盘右箭头事件
        # 其他所有情况不抑制按键事件
        return True

    def enable_space_click(self) -> bool:
        """
        启用空格键点击功能
        返回: 是否成功启用
        """
        if self._space_remapped:
            return False

        try:
            def space_handler(_):
                # 执行点击操作
                pyautogui.click(*self._space_click_position)

                # 使用区域颜色变化检测作为防抖条件
                def check_color_change():
                    # 初始颜色
                    initial_color = pyautogui.screenshot(region=(1700, 180, 20, 20)).getpixel((10, 10))

                    # 等待1秒或直到颜色变化
                    start_time = time.time()
                    while time.time() - start_time < 1:
                        current_color = pyautogui.screenshot(region=(1700, 180, 20, 20)).getpixel((10, 10))
                        if current_color != initial_color:
                            counter_manager.increment()
                            return
                        time.sleep(0.1)

                    # 1秒后仍未变化，则不调用
                    return

                # 在新线程中执行颜色检测以避免阻塞
                threading.Thread(target=check_color_change, daemon=True).start()

            # 使用on_press_key并设置suppress=True确保抑制原始事件
            keyboard.on_press_key("space", space_handler, suppress=True)
            self._space_remapped = True
            return True
        except Exception as e:
            print(f"启用空格点击失败: {e}")
            return False

    def disable_space_click(self) -> bool:
        """
        禁用空格键点击功能
        返回: 是否成功禁用
        """
        if not self._space_remapped:
            return False

        try:
            # 使用unhook_key取消空格键监听
            keyboard.unhook_key("space")
            self._space_remapped = False
            return True
        except Exception as e:
            print(f"禁用空格点击失败: {e}")
            return False

    def enable_alt_click(self) -> bool:
        """
        启用Alt键点击功能
        返回: 是否成功启用
        """
        if self._delete_listener is not None:
            return False

        try:
            self._delete_listener = keyboard.on_press_key(
                "alt",
                lambda _: self._trigger_click(),
                suppress=True
            )
            return True
        except Exception as e:
            print(f"启用删除点击失败: {e}")
            return False

    def disable_alt_click(self) -> bool:
        """
        禁用Alt键点击功能
        返回: 是否成功禁用
        """
        if self._delete_listener is None:
            return False

        try:
            keyboard.unhook(self._delete_listener)
            self._delete_listener = None
            return True
        except Exception as e:
            print(f"禁用删除点击失败: {e}")
            return False

    def _trigger_click(self) -> None:
        """执行点击操作"""
        try:
            pyautogui.click(*self._click_position)
        except Exception as e:
            print(f"点击操作失败: {e}")

    def disable_all(self) -> None:
        """禁用所有键盘映射"""
        self.disable_space_click()
        self.disable_alt_click()

# 全局单例实例
keyboard_manager = KeyboardManager()
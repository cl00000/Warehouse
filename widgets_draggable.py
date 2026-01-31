# widgets/widgets_draggable.py
from PySide6.QtCore import Qt
from PySide6.QtGui import QMouseEvent
from PySide6.QtWidgets import QWidget

"""
    ✅ 简介
        功能：添加窗口拖动能力。
        结构清晰、易于复用。
        实现自定义窗口交互的常见手段。
"""

class DraggableMixin:
    """实现窗口拖拽功能的混入类"""
    def __init__(self):
        # 初始化拖拽起点变量 _drag_start_pos 为 None，用来记录鼠标按下的位置。
        self._drag_start_pos = None

    def mousePressEvent(self, event: QMouseEvent):
        """
        当用户按下鼠标时，如果是 左键，就记录当前鼠标的全局位置（屏幕上的位置）。
        如果不是左键，就交给原始的 QWidget.mousePressEvent 处理。
        """
        if event.button() == Qt.LeftButton:
            self._drag_start_pos = event.globalPosition().toPoint()
            event.accept()
        else:
            QWidget.mousePressEvent(self, event)  # 显式调用 QWidget 的方法

    def mouseMoveEvent(self, event: QMouseEvent):
        """
        当鼠标移动时，如果是 按住左键拖动，就计算鼠标移动的偏移量 delta，并调用 self.move() 移动窗口。
        然后更新 _drag_start_pos 为当前的新位置。
        如果不是左键拖动，就仍交由 QWidget.mouseMoveEvent 处理。
        """
        if self._drag_start_pos and event.buttons() == Qt.LeftButton:
            delta = event.globalPosition().toPoint() - self._drag_start_pos
            self.move(self.pos() + delta)
            self._drag_start_pos = event.globalPosition().toPoint()
            event.accept()
        else:
            QWidget.mouseMoveEvent(self, event)  # 显式调用 QWidget 的方法

    def mouseReleaseEvent(self, event: QMouseEvent):
        """
        当释放左键时，将 _drag_start_pos 重置为 None。
        同样，如果不是左键释放，就交回给 QWidget.mouseReleaseEvent。
        """
        if event.button() == Qt.LeftButton:
            self._drag_start_pos = None
            event.accept()
        else:
            QWidget.mouseReleaseEvent(self, event)  # 显式调用 QWidget 的方法



"""
    这是一个 mixin（混入）类，用于增强现有类的功能（在这里是给窗口添加鼠标拖拽移动的能力）。它不能单独使用，必须和一个 QWidget 子类一起继承使用。
    💡 适用场景
         适用于无边框的自定义窗口（frameless window），因为这类窗口不会自动处理拖拽操作。
    
    ✅ 示例用法（简化）
            from PySide6.QtWidgets import QWidget
            from widgets.widgets_draggable import DraggableMixin
    
            class MyWindow(QWidget, DraggableMixin):
                def __init__(self):
                    super().__init__()
                    DraggableMixin.__init__(self)  # 必须显式调用
                    self.setWindowFlags(Qt.FramelessWindowHint)  # 无边框窗口
"""
# window_frosted_glass.py (Windows专用版)
from PySide6.QtWidgets import QWidget
from PySide6.QtGui import (
    QPainter,
    QColor,
    QPen,
    QBrush,
    QPainterPath,
    QPaintEvent,
)
from PySide6.QtCore import Qt
import ctypes
from typing import Optional

"""
    Windows专用毛玻璃效果窗口基类
"""

# Windows API类型定义
class ACCENTPOLICY(ctypes.Structure):
    _fields_ = [
        ("AccentState", ctypes.c_uint),
        ("AccentFlags", ctypes.c_uint),
        ("GradientColor", ctypes.c_uint),
        ("AnimationId", ctypes.c_uint),
    ]


class WINDOWCOMPOSITIONATTRIBDATA(ctypes.Structure):
    _fields_ = [
        ("Attribute", ctypes.c_int),
        ("Data", ctypes.POINTER(ctypes.c_byte)),
        ("SizeOfData", ctypes.c_size_t),
    ]


class FrostedGlassWidget(QWidget):
    """

    特性：
    - Windows亚克力效果
    - 圆角边框
    - 抗锯齿渲染
    - 可自定义样式

    使用方法：
    class MyWindow(FrostedGlassWidget):
        def __init__(self):
            super().__init__()
            # 自定义内容...
    """

    # 样式常量
    BORDER_RADIUS: int = 15
    BACKGROUND_COLOR: QColor = QColor(173, 216, 230, 170)  # 浅蓝色半透明
    BORDER_COLOR: QColor = QColor(100, 149, 237, 200)  # 矢车菊蓝半透明
    BORDER_WIDTH: int = 2

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)

        # 初始化样式对象
        self._background_brush = QBrush(self.BACKGROUND_COLOR)
        self._border_pen = QPen(self.BORDER_COLOR, self.BORDER_WIDTH)
        self._border_pen.setJoinStyle(Qt.RoundJoin)

        # 窗口属性设置
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setWindowFlags(self.windowFlags() | Qt.FramelessWindowHint)

        # 应用Windows亚克力特效
        self._apply_windows_acrylic()

    def _apply_windows_acrylic(self) -> None:
        """应用Windows亚克力特效"""
        try:
            # 获取窗口句柄
            hwnd = ctypes.windll.user32.GetParent(self.winId())

            # 设置亚克力效果参数
            accent = ACCENTPOLICY()
            accent.AccentState = 4  # ACCENT_ENABLE_ACRYLICBLURBEHIND

            # 颜色转换(ABGR格式)
            color = self.BACKGROUND_COLOR
            gradient_color = (color.alpha() << 24) | (color.red() << 16) | (color.green() << 8) | color.blue()
            accent.GradientColor = gradient_color

            # 设置窗口组合属性
            data = WINDOWCOMPOSITIONATTRIBDATA()
            data.Attribute = 19  # WCA_ACCENT_POLICY
            data.SizeOfData = ctypes.sizeof(accent)
            data.Data = ctypes.cast(ctypes.pointer(accent), ctypes.POINTER(ctypes.c_byte))

            # 调用Windows API
            ctypes.windll.user32.SetWindowCompositionAttribute(hwnd, ctypes.byref(data))
        except Exception as e:
            print(f"Windows亚克力效果应用失败: {e}")

    def paintEvent(self, event: QPaintEvent) -> None:
        """绘制窗口"""
        painter = QPainter(self)
        painter.setRenderHints(
            QPainter.Antialiasing |
            QPainter.SmoothPixmapTransform
        )

        # 创建圆角路径
        path = QPainterPath()
        path.addRoundedRect(
            self.rect().adjusted(
                self.BORDER_WIDTH / 2,
                self.BORDER_WIDTH / 2,
                -self.BORDER_WIDTH / 2,
                -self.BORDER_WIDTH / 2
            ),
            self.BORDER_RADIUS,
            self.BORDER_RADIUS
        )

        # 绘制背景(亚克力效果由Windows处理，这里只绘制半透明层)
        painter.save()
        painter.setPen(Qt.NoPen)
        painter.setBrush(self._background_brush)
        painter.drawPath(path)
        painter.restore()

        # 绘制边框
        painter.save()
        painter.setPen(self._border_pen)
        painter.setBrush(Qt.NoBrush)
        painter.drawPath(path)
        painter.restore()

        painter.end()

    def update_style(self, **kwargs) -> None:
        """
        动态更新样式

        参数:
            bg_color: 背景颜色(QColor)
            border_color: 边框颜色(QColor)
            radius: 圆角半径(int)
        """
        if 'bg_color' in kwargs:
            self.BACKGROUND_COLOR = kwargs['bg_color']
            self._background_brush = QBrush(self.BACKGROUND_COLOR)
            self._apply_windows_acrylic()  # 背景颜色变化需要重新应用亚克力效果

        if 'border_color' in kwargs:
            self.BORDER_COLOR = kwargs['border_color']
            self._border_pen.setColor(self.BORDER_COLOR)

        if 'radius' in kwargs:
            self.BORDER_RADIUS = kwargs['radius']

        self.update()
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QPushButton,
    QMessageBox, QGraphicsBlurEffect
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QPainter, QPainterPath, QPen, QColor
from widgets_draggable import DraggableMixin


"""
    🌟 总体功能
       毛玻璃效果的弹出消息框，用来代替 QMessageBox 等系统样式弹窗，用于自定义风格的应用界面
"""

class FrostedMessageBox(QDialog, DraggableMixin):
    """毛玻璃风格消息弹窗"""
    BORDER_RADIUS = 15
    BORDER_COLOR = QColor(100, 149, 237, 200)
    TITLE_STYLES = {
        QMessageBox.Warning: "color: #8B0000;",
        QMessageBox.Critical: "color: #FF4500;",
        QMessageBox.Information: "color: #2F4F4F;"
    }

    def __init__(self, parent=None, title="提示", text="", icon=QMessageBox.Information):
        super().__init__(parent)
        DraggableMixin.__init__(self)

        # 设置窗口样式：无边框、置顶、透明背景
        self.setWindowFlags(self.windowFlags() |
                            Qt.WindowStaysOnTopHint |
                            Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedSize(300, 150)

        self._init_blur_background()            # 初始化模糊背景层
        self._init_ui(title, text, icon)        # 初始化 UI 组件
        self.setCursorStyle()                   # 设置按钮鼠标样式

    def _init_blur_background(self):
        """设置毛玻璃背景层"""
        self._blur_background = QLabel(self)
        self._blur_background.setGeometry(self.rect())
        self._blur_background.lower()  # 放到底层

        blur_effect = QGraphicsBlurEffect()
        blur_effect.setBlurRadius(15)
        self._blur_background.setGraphicsEffect(blur_effect)

        self._blur_background.setStyleSheet(f"""
            background-color: rgba(173, 216, 230, 220);  /* 淡蓝半透明 */
            border-radius: {self.BORDER_RADIUS}px;
        """)

    def _init_ui(self, title, text, icon):
        """初始化内容布局"""
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        title_label = QLabel(title)
        title_label.setFont(QFont("Microsoft YaHei", 12, QFont.Bold))
        title_label.setStyleSheet(self.TITLE_STYLES.get(icon, ""))

        content_label = QLabel(text)
        content_label.setFont(QFont("Microsoft YaHei", 10))
        content_label.setStyleSheet("color: #2F4F4F;")
        content_label.setWordWrap(True)

        confirm_btn = self._create_button("确定", self.accept)

        layout.addWidget(title_label)
        layout.addWidget(content_label)
        layout.addWidget(confirm_btn, alignment=Qt.AlignCenter)
        self.setLayout(layout)

    def _create_button(self, text, callback):
        """创建带样式的按钮"""
        btn = QPushButton(text)
        btn.setFixedSize(80, 30)
        btn.setStyleSheet("""
            QPushButton {
                background: rgba(100, 149, 237, 150);
                border: 1px solid #4682B4;
                border-radius: 8px;
                color: white;
            }
            QPushButton:hover {
                background: rgba(70, 130, 180, 200);
            }
        """)
        btn.clicked.connect(callback)
        return btn

    def setCursorStyle(self):
        """为所有按钮设置手型光标"""
        for btn in self.findChildren(QPushButton):
            btn.setCursor(Qt.PointingHandCursor)

    def resizeEvent(self, event):
        """窗口尺寸变化时同步模糊背景尺寸"""
        super().resizeEvent(event)
        self._blur_background.setGeometry(self.rect())

    def paintEvent(self, event):
        """绘制圆角边框"""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        path = QPainterPath()
        path.addRoundedRect(self.rect(), self.BORDER_RADIUS, self.BORDER_RADIUS)

        painter.setPen(QPen(self.BORDER_COLOR, 2))
        painter.setBrush(Qt.NoBrush)
        painter.drawPath(path)

    # 拖拽支持
    # 覆盖了以下鼠标事件并显式调用DraggableMixin的方法，使窗口可以被鼠标左键拖动：
    def mousePressEvent(self, event):
        DraggableMixin.mousePressEvent(self, event)

    def mouseMoveEvent(self, event):
        DraggableMixin.mouseMoveEvent(self, event)

    def mouseReleaseEvent(self, event):
        DraggableMixin.mouseReleaseEvent(self, event)


"""
        ✅ 关键特性详解
            继承自 QDialog 和 DraggableMixin，具备对话框功能和拖动功能。
            
"""
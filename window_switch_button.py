from PySide6.QtWidgets import QWidget, QApplication
from PySide6.QtCore import (
    Qt, QPropertyAnimation, QEasingCurve,
    Property, QPoint, Signal, QRect
)
from PySide6.QtGui import (
    QPainter, QColor, QPen, QBrush, QPaintEvent, QResizeEvent, QKeyEvent
)
from typing import Optional, Union, Tuple


class SwitchButton(QWidget):
    """
    自定义开关按钮控件 (优化版)

    特性：
    - 平滑的滑动动画(支持自定义时长和缓动曲线)
    - 完整的颜色自定义能力
    - 丰富的状态变化信号
    - 支持直接设置状态(可选择是否使用动画)
    - 键盘和鼠标操作支持
    - 禁用状态可视化
    - 高性能绘制优化

    信号：
    - stateChanged(checked: bool): 当开关状态改变时触发
    - sliderPositionChanged(pos: QPoint): 滑块位置变化时触发

    用法示例：
        switch = SwitchButton()
        switch.setChecked(True)  # 设置开启状态(带动画)
        switch.toggle()  # 切换状态
        switch.setColors(bg_on=QColor(255,0,0))  # 自定义颜色
    """

    class Dimensions:
        """控件尺寸常量"""
        WIDTH = 60
        HEIGHT = 30
        SLIDER_SIZE = 24
        PADDING = 3
        BORDER_RADIUS = 15
        BORDER_WIDTH = 1
        ANIMATION_DURATION = 100  # 毫秒

    class DefaultColors:
        """默认颜色配置"""
        BG_OFF = QColor(200, 220, 240)  # 关闭状态背景
        BG_ON = QColor(70, 130, 180)  # 开启状态背景
        SLIDER = QColor(255, 255, 255)  # 滑块颜色
        BORDER = QColor(150, 180, 210)  # 边框颜色
        DISABLED = QColor(200, 200, 200)  # 禁用状态颜色

    # 信号定义
    stateChanged = Signal(bool)  # 状态变化信号
    sliderPositionChanged = Signal(QPoint)  # 滑块位置变化信号

    def __init__(self, parent: Optional[QWidget] = None):
        """
        初始化开关按钮控件

        参数:
            parent: 父级控件 (可选)
        """
        super().__init__(parent)
        self._setup_ui()
        self._init_animation()

    def _setup_ui(self) -> None:
        """初始化UI配置"""
        # 设置固定尺寸
        self.setFixedSize(self.Dimensions.WIDTH, self.Dimensions.HEIGHT)
        self.setFocusPolicy(Qt.StrongFocus)  # 支持键盘焦点

        # 初始化状态
        self._checked = False
        self._slider_pos = QPoint(self.Dimensions.PADDING, self.Dimensions.PADDING)

        # 初始化颜色
        self._bg_color_off = self.DefaultColors.BG_OFF
        self._bg_color_on = self.DefaultColors.BG_ON
        self._slider_color = self.DefaultColors.SLIDER
        self._border_color = self.DefaultColors.BORDER
        self._disabled_color = self.DefaultColors.DISABLED

        # 预创建绘制对象
        self._border_pen = QPen(self._border_color, self.Dimensions.BORDER_WIDTH)
        self._slider_brush = QBrush(self._slider_color)
        self._disabled_brush = QBrush(self._disabled_color)

    def _init_animation(self) -> None:
        """初始化动画系统"""
        self.animation = QPropertyAnimation(self, b"slider_pos")
        self.animation.setDuration(self.Dimensions.ANIMATION_DURATION)
        self.animation.setEasingCurve(QEasingCurve.Linear)  # 线性动画
        self.animation.finished.connect(self._animation_finished)

    def _animation_finished(self) -> None:
        """动画完成时的处理"""
        self.update()
        self.sliderPositionChanged.emit(self._slider_pos)

    def setChecked(self, checked: bool, animate: bool = True) -> None:
        """
        设置开关状态

        参数:
            checked: 是否选中
            animate: 是否使用动画效果 (默认True)
        """
        if self._checked == checked:
            return  # 状态未变化直接返回

        self._checked = checked
        target_x = self._calculate_target_x()

        if animate:
            self._start_animation(target_x)
        else:
            self._slider_pos = QPoint(target_x, self.Dimensions.PADDING)
            self.update()

        self.stateChanged.emit(self._checked)

    def _calculate_target_x(self) -> int:
        """计算滑块目标X坐标"""
        return (self.width() - self.Dimensions.SLIDER_SIZE - self.Dimensions.PADDING
                if self._checked else self.Dimensions.PADDING)

    def _start_animation(self, target_x: int) -> None:
        """启动滑动动画"""
        self.animation.stop()  # 停止当前动画
        self.animation.setStartValue(self._slider_pos)
        self.animation.setEndValue(QPoint(target_x, self.Dimensions.PADDING))
        self.animation.start()

    def toggle(self, animate: bool = True) -> None:
        """
        切换开关状态

        参数:
            animate: 是否使用动画效果 (默认True)
        """
        self.setChecked(not self._checked, animate)

    def isChecked(self) -> bool:
        """获取当前开关状态"""
        return self._checked

    # 属性动画相关方法
    def get_slider_pos(self) -> QPoint:
        """获取滑块当前位置 (用于属性动画)"""
        return self._slider_pos

    def set_slider_pos(self, pos: QPoint) -> None:
        """设置滑块位置 (用于属性动画)"""
        # 边界检查确保滑块不会超出范围
        x = min(max(pos.x(), self.Dimensions.PADDING),
                self.width() - self.Dimensions.SLIDER_SIZE - self.Dimensions.PADDING)
        y = min(max(pos.y(), self.Dimensions.PADDING),
                self.height() - self.Dimensions.SLIDER_SIZE - self.Dimensions.PADDING)
        self._slider_pos = QPoint(x, y)
        self.update()
        self.sliderPositionChanged.emit(self._slider_pos)

    slider_pos = Property(QPoint, get_slider_pos, set_slider_pos)

    def paintEvent(self, event: QPaintEvent) -> None:
        """
        绘制开关控件

        参数:
            event: 绘制事件
        """
        painter = QPainter(self)
        try:
            # 启用抗锯齿和高质量渲染
            painter.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform)

            # 计算背景区域(考虑边框宽度)
            bg_rect = QRect(
                self.Dimensions.BORDER_WIDTH,
                self.Dimensions.BORDER_WIDTH,
                self.width() - 2 * self.Dimensions.BORDER_WIDTH,
                self.height() - 2 * self.Dimensions.BORDER_WIDTH
            )

            # 绘制背景
            bg_color = self._bg_color_on if self._checked else self._bg_color_off
            if not self.isEnabled():
                bg_color = self._disabled_color

            painter.setPen(self._border_pen)
            painter.setBrush(QBrush(bg_color))
            painter.drawRoundedRect(bg_rect,
                                    self.Dimensions.BORDER_RADIUS,
                                    self.Dimensions.BORDER_RADIUS)

            # 绘制滑块
            slider_brush = self._slider_brush if self.isEnabled() else self._disabled_brush
            painter.setPen(Qt.NoPen)
            painter.setBrush(slider_brush)
            painter.drawEllipse(
                self._slider_pos.x(),
                self._slider_pos.y(),
                self.Dimensions.SLIDER_SIZE,
                self.Dimensions.SLIDER_SIZE
            )
        finally:
            painter.end()

    def mousePressEvent(self, event) -> None:
        """
        鼠标点击事件处理

        参数:
            event: 鼠标事件
        """
        if event.button() == Qt.LeftButton and self.isEnabled():
            self.toggle()
            event.accept()
        else:
            super().mousePressEvent(event)

    def keyPressEvent(self, event: QKeyEvent) -> None:
        """
        键盘事件处理

        参数:
            event: 键盘事件
        """
        if event.key() in (Qt.Key_Space, Qt.Key_Enter, Qt.Key_Return) and self.isEnabled():
            self.toggle()
            event.accept()
        else:
            super().keyPressEvent(event)

    def resizeEvent(self, event: QResizeEvent) -> None:
        """
        尺寸变化事件处理

        参数:
            event: 尺寸事件
        """
        super().resizeEvent(event)
        # 尺寸变化时重新定位滑块
        self._slider_pos = QPoint(
            self._calculate_target_x() if self._checked else self.Dimensions.PADDING,
            self.Dimensions.PADDING
        )

    def setColors(
            self,
            bg_on: Optional[Union[QColor, Tuple[int, int, int]]] = None,
            bg_off: Optional[Union[QColor, Tuple[int, int, int]]] = None,
            slider: Optional[Union[QColor, Tuple[int, int, int]]] = None,
            border: Optional[Union[QColor, Tuple[int, int, int]]] = None,
            disabled: Optional[Union[QColor, Tuple[int, int, int]]] = None
    ) -> None:
        """
        设置控件颜色配置

        参数:
            bg_on: 开启状态背景色 (QColor或RGB元组)
            bg_off: 关闭状态背景色 (QColor或RGB元组)
            slider: 滑块颜色 (QColor或RGB元组)
            border: 边框颜色 (QColor或RGB元组)
            disabled: 禁用状态颜色 (QColor或RGB元组)

        示例:
            switch.setColors(
                bg_on=(255, 0, 0),  # 红色开启状态
                slider=(255, 255, 0)  # 黄色滑块
            )
        """
        if bg_on is not None:
            self._bg_color_on = self._parse_color(bg_on)
        if bg_off is not None:
            self._bg_color_off = self._parse_color(bg_off)
        if slider is not None:
            self._slider_color = self._parse_color(slider)
            self._slider_brush = QBrush(self._slider_color)
        if border is not None:
            self._border_color = self._parse_color(border)
            self._border_pen.setColor(self._border_color)
        if disabled is not None:
            self._disabled_color = self._parse_color(disabled)
            self._disabled_brush = QBrush(self._disabled_color)

        self.update()

    @staticmethod
    def _parse_color(color: Union[QColor, Tuple[int, int, int]]) -> QColor:
        """将颜色参数转换为QColor对象"""
        if isinstance(color, QColor):
            return color
        elif isinstance(color, (tuple, list)) and len(color) == 3:
            return QColor(*color)
        else:
            raise ValueError("颜色参数必须是QColor或RGB元组")

    @Property(bool)
    def checked(self) -> bool:
        """获取或设置开关状态(支持动画)"""
        return self._checked

    @checked.setter
    def checked(self, value: bool):
        self.setChecked(value)

    def __repr__(self) -> str:
        """返回控件的字符串表示(用于调试)"""
        return (f"<SwitchButton at {hex(id(self))}, "
                f"checked={self._checked}, "
                f"position={self._slider_pos}, "
                f"enabled={self.isEnabled()}>")


# 示例用法
if __name__ == "__main__":
    app = QApplication([])

    switch = SwitchButton()
    switch.setColors(
        bg_on=(0, 255, 0),  # 绿色开启状态
        bg_off=(255, 0, 0),  # 红色关闭状态
        slider=(255, 255, 255)
    )
    switch.show()

    app.exec()
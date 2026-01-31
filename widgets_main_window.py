# widgets_main_window.py
"""
主窗口模块 - 实现带毛玻璃效果的现代化主界面

包含功能：
- 可拖拽的毛玻璃效果窗口
- 多个功能开关控制
- 坐标输入验证
- Excel对比功能
- 系统托盘控制
"""
import time

from function_counter import counter_manager
from window_frosted_glass import FrostedGlassWidget
from window_switch_button import SwitchButton
from function_checkColor import start_monitoring, stop_monitoring
import function_switch2
from function_keyboard_manager import keyboard_manager
from function_config_manager import load_config, save_config, update_ocr_config
from function_group_calculation import group_calculation
import re
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QLineEdit, \
    QMessageBox
from PySide6.QtCore import Qt, Slot, QTimer
from PySide6.QtGui import QFont
from widgets_draggable import DraggableMixin
from widgets_frosted_message_box import FrostedMessageBox
from function_create_auxiliary_table import create_auxiliary_table


def _create_stat_label(text):
    """创建统计信息标签"""
    label = QLabel(text)
    label.setFont(QFont("Microsoft YaHei", 9))
    label.setStyleSheet("""
        color: #2F4F4F;
        background: rgba(255, 255, 255, 80);
        border-radius: 5px;
        padding: 3px 8px;
    """)
    label.setAlignment(Qt.AlignCenter)
    return label


class ModernWindow(DraggableMixin, FrostedGlassWidget):
    """
    现代化主窗口类，继承可拖拽和毛玻璃效果特性

    特性：
    - 可拖拽的标题栏区域
    - 毛玻璃透明效果
    - 三个功能切换开关
    - 坐标输入验证
    - 创建副表功能
    - 系统托盘控制
    """

    # 主窗体尺寸常量
    WINDOW_SIZE = (260, 388)

    # 样式表常量
    STYLE_TITLE = "color: #2F4F4F;"
    # 坐标输入框样式
    STYLE_INPUT = """
        QLineEdit {
            background: rgba(255, 255, 255, 150);
            border: 1px solid #4682B4;
            border-radius: 10px;
            padding: 5px;
            height:30px; 
            color: #2F4F4F;
        }
        QLineEdit:focus {
            border: 1px solid #1E90FF;
            color: #000000;
        }
        QLineEdit::placeholder {
            color: #708090;
        }
    """

    # 按钮样式表
    BUTTON_STYLE = {
        "min": """
            QPushButton { 
                background-color: #FFBD2E; 
                border-radius: 6px; 
            }
            QPushButton:hover { 
                background-color: #FF9F00; 
            }
        """,
        "close": """
            QPushButton { 
                background-color: #FF5F56; 
                border-radius: 6px; 
            }
            QPushButton:hover { 
                background-color: #FF3B30; 
            }
        """,
        "action": """
            QPushButton { 
                background-color: #5F9EA0; 
                border-radius: 10px;
                padding: 8px;
                color: white;
            }
            QPushButton:hover { 
                background-color: #4682B4;
            }
        """
    }

    def __init__(self):
        FrostedGlassWidget.__init__(self)
        DraggableMixin.__init__(self)

        self.config = load_config()
        self.switch_buttons = []

        self._setup_window_properties()
        self._init_ui()  # 先初始化UI

        # 后加载状态和位置
        self._load_switch_states()
        self._load_window_position()  # 位置加载应该在方法内完成所有逻辑

        self._init_refresh_timer()

        # 连接键盘信号
        keyboard_manager.left_key_pressed.connect(self._handle_left_key)
        keyboard_manager.right_key_pressed.connect(self._handle_right_key)

    def _init_refresh_timer(self):
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_counts)
        self.timer.start(1000)  # 每秒更新一次

    def update_counts(self):
        today, session = counter_manager.get_counts()
        self.today_label.setText(f"今日录入数量：{today}")
        self.session_label.setText(f"本次录入数量：{session}")

        # 计算并显示录入速度
        speed = counter_manager.calculate_speed()
        if speed is None:
            self.speed_label.setText("平均速度：--")
        else:
            self.speed_label.setText(f"平均速度：{speed:.1f}秒/单")



    @Slot()
    def _handle_left_key(self):
        """处理左方向键事件 - 关闭当前开启的OCR或自动入库开关"""
        # 如果自动入库开关开启，则关闭它
        if self.auto_switch.isChecked():
            self.auto_switch.setChecked(False, animate=True)
        # 如果OCR开关开启，则关闭它
        elif self.ocr_switch.isChecked():
            self.ocr_switch.setChecked(False, animate=True)

    @Slot()
    def _handle_right_key(self):
        """处理右方向键事件"""
        if not self.ocr_switch.isChecked() and not self.auto_switch.isChecked():
            self.ocr_switch.setChecked(True, animate=True)
        else:
            current_state = self.ocr_switch.isChecked()
            self.ocr_switch.setChecked(not current_state, animate=True)
            self.auto_switch.setChecked(current_state, animate=True)

    def _load_switch_states(self):
        """从配置文件加载所有开关状态"""
        config = load_config()
        for i, switch in enumerate(self.switch_buttons, start=1):
            state = config.get(f"switch{i}_state", False)
            switch.setChecked(state, animate=False)

    def _save_switch_state(self, checked: bool, switch_num: int):
        """保存指定开关的状态到配置文件"""
        config = load_config()
        config[f"switch{switch_num}_state"] = checked
        save_config(config)

    def _load_window_position(self):
        """完整的位置加载方法"""
        try:
            config = load_config()
            default_pos = [200, 200]
            pos = config.get("window_position", default_pos)

            # 获取屏幕信息
            screen = QApplication.primaryScreen()
            if not screen:
                self.move(*default_pos)
                return

            screen_geo = screen.availableGeometry()

            # 计算安全位置
            max_x = screen_geo.right() - self.width()
            max_y = screen_geo.bottom() - self.height()
            x = min(max(pos[0], screen_geo.left()), max_x)
            y = min(max(pos[1], screen_geo.top()), max_y)

            self.move(int(x), int(y))
        except Exception as e:
            print(f"加载窗口位置失败: {str(e)}")
            self.move(100, 100)  # 失败时使用默认位置

    def _save_window_position(self):
        """保存当前窗口位置到配置文件"""
        pos = [self.x(), self.y()]
        config = load_config()
        config["window_position"] = pos
        save_config(config)

    # ------------------------ 初始化方法 ------------------------
    def _setup_window_properties(self):
        """设置窗口基本属性"""
        self.setWindowTitle("仓库工具")
        self.setFixedSize(*self.WINDOW_SIZE)
        self.setWindowFlags(
            self.windowFlags() |
            Qt.WindowStaysOnTopHint |
            Qt.FramelessWindowHint
        )

    def _init_ui(self):
        """初始化用户界面"""
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(2)  # 垂直间距

        # 按层次添加界面组件
        main_layout.addLayout(self._create_top_bar())  # 顶部控制栏
        self._create_switches(main_layout)  # 功能开关组
        main_layout.addStretch()  # 弹性空间
        main_layout.addLayout(self._create_action_button())  # 功能按钮
        main_layout.addLayout(self._create_bottom_input())  # 底部输入区

        self.setLayout(main_layout)

    # ------------------------ UI组件创建方法 ------------------------
    def _create_top_bar(self):
        """创建顶部控制栏布局"""
        top_bar = QHBoxLayout()
        top_bar.setContentsMargins(0, 0, 0, 15)

        # 标题标签（居中显示）
        title = self._create_label("仓库工具", font_size=12, bold=True)
        top_bar.addStretch(1)
        top_bar.addWidget(title, alignment=Qt.AlignCenter)
        top_bar.addStretch(1)

        # 窗口控制按钮（最小化/关闭）
        control_buttons = QHBoxLayout()
        control_buttons.addWidget(
            self._create_control_button("min", self.showMinimized)
        )
        control_buttons.addWidget(
            self._create_control_button("close", self.close)
        )

        top_bar.addLayout(control_buttons)
        return top_bar

    def _create_control_button(self, btn_type, callback):
        """
        创建通用控制按钮

        :param btn_type: 按钮类型（min/close）
        :param callback: 点击回调函数
        """
        btn = QPushButton()
        btn.setFixedSize(12, 12)
        btn.setStyleSheet(self.BUTTON_STYLE[btn_type])
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.clicked.connect(callback)
        return btn

    def _create_switches(self, layout):
        """创建功能开关组（带状态持久化）"""
        switch_configs = [
            ("OCR自动入库", self.on_switch1),
            ("自动入库", self.on_switch2),
            ("按键映射", self.on_switch3)
        ]

        for idx, (text, callback) in enumerate(switch_configs):
            switch_layout = QHBoxLayout()
            switch_layout.setContentsMargins(0, 0, 0, 0)
            switch_layout.setSpacing(0)

            # 开关标签
            label = self._create_label(text, font_size=10)

            # 开关容器
            container = QWidget()
            container.setFixedSize(80, 40)
            container.setStyleSheet("""
                QWidget { background: transparent; }
                QWidget:hover { 
                    background: rgba(100, 155, 143, 60); 
                    border-radius: 5px; 
                }
            """)
            container.setCursor(Qt.PointingHandCursor)

            # 开关按钮本体
            switch = SwitchButton(container)
            switch.move(10, 5)
            self.switch_buttons.append(switch)

            # 开关连接状态保存信号
            if idx < 3:  # 0=OCR, 1=自动入库, 2=按键映射
                switch.stateChanged.connect(
                    lambda checked, switch_idx=idx: self._save_switch_state(checked, switch_idx + 1)
                )

            switch.stateChanged.connect(callback)  # 原有业务逻辑不变

            switch_layout.addWidget(label)
            switch_layout.addStretch()
            switch_layout.addWidget(container)
            layout.addLayout(switch_layout)

            if idx == 2:  # 在第三个开关后添加统计信息
                stats_layout = QVBoxLayout()
                stats_layout.setContentsMargins(5, 5, 5, 5)
                stats_layout.setSpacing(5)

                # 今日录入数量
                self.today_label = _create_stat_label("今日录入数量：0")
                # 本次录入数量
                self.session_label = _create_stat_label("本次录入数量：0")
                # 录入速度
                self.speed_label = _create_stat_label("平均速度：--")

                stats_layout.addWidget(self.today_label)
                stats_layout.addWidget(self.session_label)
                stats_layout.addWidget(self.speed_label)
                layout.addLayout(stats_layout)

            layout.addLayout(switch_layout)

        # 保存关键开关引用
        self.ocr_switch, self.auto_switch = self.switch_buttons[:2]

        # 仅加载前3个开关的状态
        self._load_switch_states()

        # 初始化速度计算相关变量
        self.last_count = 0
        self.last_time = time.time()

    def _create_action_button(self):
        """创建功能按钮布局"""
        button_layout = QHBoxLayout()
        button_layout.setContentsMargins(0, 10, 0, 10)

        create_table_btn = QPushButton("创建副表")
        create_table_btn.setStyleSheet(self.BUTTON_STYLE["action"])
        create_table_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        create_table_btn.clicked.connect(self.on_create_auxiliary_table)

        excel_btn_new = QPushButton("分组计算")
        excel_btn_new.setStyleSheet(self.BUTTON_STYLE["action"])
        excel_btn_new.setCursor(Qt.CursorShape.PointingHandCursor)
        excel_btn_new.clicked.connect(self.on_group_calculation)

        button_layout.addWidget(create_table_btn)  # 使用新按钮
        button_layout.addWidget(excel_btn_new)
        return button_layout

    def _create_bottom_input(self):
        """创建底部输入区域"""
        bottom_layout = QVBoxLayout()
        bottom_layout.setSpacing(8)

        # 坐标输入框
        self.text_input = QLineEdit()
        self.text_input.setPlaceholderText("请输入坐标：#xxx,xxx#")
        self.text_input.setStyleSheet(self.STYLE_INPUT)
        self.text_input.returnPressed.connect(self._handle_input)

        bottom_layout.addWidget(self.text_input)
        return bottom_layout

    def _create_label(self, text, font_size=10, bold=False, color=None):
        """
        创建标准化标签

        :param text: 显示文本
        :param font_size: 字体大小
        :param bold: 是否加粗
        :param color: 字体颜色
        """
        label = QLabel(text)
        label.setFont(QFont("Microsoft YaHei", font_size,
                            QFont.Bold if bold else QFont.Normal))
        style = "color: #2F4F4F !important;"  # 使用!important覆盖继承样式
        if color:
            style = f"color: {color} !important;"
        label.setStyleSheet(style)
        return label

    # ------------------------ 业务逻辑方法 ------------------------
    def _handle_input(self):
        """处理坐标输入"""
        input_text = self.text_input.text().strip()

        # 统一处理中文逗号
        normalized_text = input_text.replace('，', ',')

        # 正则验证输入格式
        if not re.match(r"^#\s*(\d+)\s*[,，]\s*(\d+)\s*#", normalized_text):
            self.show_message("格式错误", "请输入正确格式：#xxx,xxx#\n例如：#123,456#")
            return

        try:
            # 提取坐标值
            x1, x2 = map(int, re.findall(r"\d+", normalized_text))

            # 坐标范围验证
            if not (0 <= x1 <= 1920 and 0 <= x2 <= 1920):
                raise ValueError("坐标超出有效范围(0-1920)")

            # 更新配置
            self.config.update({"region1_x": x1, "region2_x": x2})
            save_config(self.config)
            update_ocr_config(x1, x2)

            self.show_message("成功", f"坐标已更新：({x1}, {x2})", QMessageBox.Information)
            self.text_input.clear()

        except Exception as e:
            self.show_message("错误", f"保存失败：{str(e)}", QMessageBox.Critical)

    def show_message(self, title, text, icon=QMessageBox.Warning):
        """显示毛玻璃风格消息弹窗"""
        FrostedMessageBox(self, title, text, icon).exec()

    # ------------------------ 事件处理槽函数 ------------------------
    @Slot(bool)
    def on_switch1(self, checked):
        """OCR自动入库开关回调"""
        # 互斥逻辑：开启OCR时关闭自动入库
        if checked and self.auto_switch.isChecked():
            self.auto_switch.setChecked(False)

        try:
            start_monitoring() if checked else stop_monitoring()
        except Exception as e:
            self.show_message("错误", f"OCR监控异常: {str(e)}")

    @Slot(bool)
    def on_switch2(self, checked):
        """自动入库开关回调"""
        # 互斥逻辑：开启自动入库时关闭OCR
        if checked and self.ocr_switch.isChecked():
            self.ocr_switch.setChecked(False)

        function_switch2.start_monitoring() if checked else function_switch2.stop_monitoring()

    @Slot(bool)
    def on_switch3(self, checked):
        """修改后的按键映射开关回调"""
        if checked:
            keyboard_manager.enable_keyboard_mapping()  # 统一启用所有键盘功能
        else:
            keyboard_manager.disable_keyboard_mapping()  # 统一禁用所有键盘功能

    @Slot()
    def on_create_auxiliary_table(self):
        """创建副表按钮点击回调"""
        success, message = create_auxiliary_table()
        if not success and message:  # 只有出错时才显示消息
            self.show_message("警告", message, QMessageBox.Warning)

    @Slot()
    def on_group_calculation(self):
        """分组计算按钮点击回调"""
        try:
            success, message = group_calculation()
            if not success:
                # 如果有错误消息，显示弹窗
                self.show_message("警告", message, QMessageBox.Warning)
        except Exception as e:
            self.show_message("错误", f"分组计算执行出错: {str(e)}", QMessageBox.Critical)

    # ------------------------ 辅助方法 ------------------------
    def _safe_stop(self, stop_func, name):
        """
        安全停止功能

        :param stop_func: 停止方法
        :param name: 功能名称（用于错误提示）
        """
        try:
            stop_func()
        except Exception as e:
            print(f"停止 {name} 时出错: {e}")

    def closeEvent(self, event):
        # 保存窗口位置前先更新一次计数（确保最新状态被保存）
        counter_manager.get_counts()  # 这会触发条件检查并保存
        self._save_window_position()
        """优化关闭事件处理"""
        # 保存状态前检查窗口是否可见
        if self.isVisible():
            self._save_window_position()

        # 使用线程安全的方式停止功能
        self._stop_all_services()

        super().closeEvent(event)

    def _stop_all_services(self):
        """集中停止所有服务"""
        services = [
            (function_switch2.stop_monitoring, "自动入库"),
            (stop_monitoring, "OCR监控"),
            (lambda: keyboard_manager.disable_all(), "键盘映射")
        ]

        for func, name in services:
            self._safe_stop(func, name)
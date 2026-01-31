import sys
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication
from widgets_main_window import ModernWindow

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    # 设置应用程序图标（影响任务栏和所有窗口）
    app.setWindowIcon(QIcon("fox.ico"))

    window = ModernWindow()
    window.show()
    sys.exit(app.exec())
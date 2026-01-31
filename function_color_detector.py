#function/function_color_detector.py

from PIL import ImageGrab
import cv2
import numpy as np

"""
    颜色检测功能
    检测指定区域的颜色并返回结果
"""
def get_color_at_region(x=1700, y=180, width=20, height=20):
    """获取屏幕指定区域的平均颜色（默认 1700,180 20x20）"""
    screenshot = ImageGrab.grab(bbox=(x, y, x + width, y + height))
    screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
    average_color = np.mean(screenshot, axis=(0, 1))
    return average_color[::-1]  # 转换为 RGB

def detect_color(color):
    """根据 HSV 颜色空间判断颜色类型
    返回: 'green', 'red', 或 'unknown'"""
    color_hsv = cv2.cvtColor(np.uint8([[color]]), cv2.COLOR_RGB2HSV)[0][0]
    hue = color_hsv[0]

    if 35 <= hue <= 85:  # 绿色范围
        return "green"
    elif hue <= 10 or hue >= 170:  # 红色范围
        return "red"
    else:
        return "unknown"
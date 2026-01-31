# import threading
# import time
# import numpy as np
# import winsound
# import mss
# from PIL import Image
# from PySide6.QtCore import QThread, Signal, QMutex
# from function_config_manager import load_config
# from function_OCR import ocr_parallel_scan  # 直接调用 OCR 主流程
#
#
# class NotesDetector(QThread):
#     detection_signal = Signal(str)
#
#     def __init__(self):
#         super().__init__()
#         self.mutex = QMutex()
#         self.running = False
#         self.last_hash = None
#         self.check_interval = 0.5
#
#     def run(self):
#         self.running = True
#         print("[备注检测] 线程已启动")
#         print(f"[备注检测] 开始监测备注区域，检测间隔: {self.check_interval}秒")
#
#         while self.running:
#             try:
#                 config = load_config()
#                 x, y, w, h = config["notes_position"]
#                 # print(f"[Debug] 当前备注区域坐标: x={x}, y={y}, w={w}, h={h}")
#
#                 # 截取当前区域
#                 with mss.mss() as sct:
#                     screenshot = sct.grab({"left": x, "top": y, "width": w, "height": h})
#                     img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")
#
#                 # 计算图像哈希
#                 current_hash = self.image_hash(img)
#                 # print(f"[备注检测] 当前区域哈希值: {current_hash[:10]}...")
#
#                 # 检测到变化时处理
#                 if current_hash != self.last_hash:
#                     print("⚠️ [备注检测] 检测到像素变化！开始OCR验证...")
#                     self.last_hash = current_hash
#                     self.process_notes_change()
#
#             except Exception as e:
#                 print(f"[备注检测] 异常: {str(e)}")
#
#             time.sleep(self.check_interval)
#
#         print("[备注检测] 线程已停止")
#
#     def image_hash(self, image):
#         """计算图像感知哈希"""
#         resized = image.resize((8, 8), Image.LANCZOS).convert("L")
#         pixels = np.array(resized)
#         avg = np.mean(pixels)
#         return bytes((pixels > avg).flatten())
#
#     def process_notes_change(self):
#         """调用 ocr_parallel_scan 进行验证"""
#         try:
#             results, _ = ocr_parallel_scan()
#             notes_result = next(
#                 (item for item in results if item["region"] == "区域4(12345)"),
#                 {"result": "未找到备注区域"}
#             )
#
#             if notes_result["result"] == "12345":
#                 print("🔔 [备注检测] 检测到有效备注内容: 12345")
#                 # 创建独立线程播放声音
#                 sound_thread = threading.Thread(target=play_alert_sound)
#                 sound_thread.start()
#                 # 发送信号（确保UI操作在主线程）
#                 self.detection_signal.emit("12345")
#
#         except Exception as e:
#             print(f"[备注检测] OCR验证失败: {str(e)}")
#
#     def stop(self):
#         """安全停止检测"""
#         self.mutex.lock()
#         self.running = False
#         self.mutex.unlock()
#         # print("[备注检测] 正在停止线程...")
#
#
# def play_alert_sound():
#     """播放提示音"""
#     print("[备注检测] 播放提示音...")
#     for _ in range(3):
#         winsound.Beep(4000, 150)
#         time.sleep(0.1)

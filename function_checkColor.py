import pyautogui
import winsound
from function_OCR import ocr_parallel_scan
import threading
import time
from function_color_detector import get_color_at_region, detect_color
from typing import Optional
from function_counter import counter_manager

# 全局线程控制
monitor_thread: Optional[threading.Thread] = None
stop_event = threading.Event()
lock = threading.Lock()

# 屏幕区域坐标
REGION2 = (1700, 280, 20, 20)  # 第二行检测区域
REGION3 = (1700, 380, 20, 20)  # 第三行检测区域
CLICK_POSITION = (650, 983)  # 点击位置：退货按钮


def stop_monitoring():
    """安全停止监控线程"""
    global monitor_thread
    with lock:
        if monitor_thread and monitor_thread.is_alive():
            stop_event.set()
            monitor_thread.join(timeout=1)
            print("OCR监控已安全停止")
        monitor_thread = None
        stop_event.clear()


def start_monitoring():
    """启动监控线程"""
    global monitor_thread
    with lock:
        if monitor_thread and monitor_thread.is_alive():
            print("OCR监控已在运行中")
            return

        monitor_thread = threading.Thread(
            target=monitor_color_changes,
            daemon=True
        )
        stop_event.clear()
        monitor_thread.start()
        print("OCR监控已启动")


def is_region_white(region):
    """检测指定区域是否为纯白色（RGB值均>240）"""
    screenshot = pyautogui.screenshot(region=region)
    return all(p[0] > 240 and p[1] > 240 and p[2] > 240 for p in screenshot.getdata())


def wait_until_green_disappears(timeout=600):
    """等待绿色区域消失，带超时机制(10分钟)"""
    start_time = time.time()
    while (time.time() - start_time) < timeout and not stop_event.is_set():
        if detect_color(get_color_at_region()) != "green":
            return True
        time.sleep(0.1)
    return False


def click_return_and_wait():
    """点击退货按钮并等待绿色消失"""
    pyautogui.click(*CLICK_POSITION)
    counter_manager.increment()
    wait_until_green_disappears()


def monitor_color_changes():
    """优化后的核心检测逻辑"""
    try:
        while not stop_event.is_set():
            if detect_color(get_color_at_region()) == "green":
                print("\n--- 检测到绿色，开始处理流程 ---")
                process_start = time.time()

                # 第一步：只扫描必要的基础区域
                results_dict, _ = ocr_parallel_scan(['digit', 'letter', 'NewProduct'])  # 解包元组
                if stop_event.is_set():
                    break

                # 提取基础信息
                quantity_valid = results_dict.get('digit', '') == '1'  # 现在操作的是字典
                product_code = results_dict.get('letter', '')
                product_code2 = results_dict.get('NewProduct', '')

                if quantity_valid:
                    if product_code == 'ABCD':
                        region2_white = is_region_white(REGION2)

                        if region2_white:
                            # 情况1：无赠品
                            play_sound("success")
                            click_return_and_wait()
                            print(f'✅整套-无赠品 (耗时: {time.time() - process_start:.2f}s)')
                        else:
                            # 第二步：需要时再扫描赠品区域
                            gift_results, _ = ocr_parallel_scan(['extended'])  # 解包元组
                            gift_code = gift_results.get('extended', '')

                            if gift_code == 'ABCD':
                                region3_white = is_region_white(REGION3)

                                if region3_white:
                                    play_sound("success")
                                    click_return_and_wait()
                                    print(f'✅整套-有赠品 (耗时: {time.time() - process_start:.2f}s)')
                                else:
                                    play_sound("failure")
                                    wait_until_green_disappears()
                                    print(f'❌整套有赠品，但有多件产品 (耗时: {time.time() - process_start:.2f}s)')
                            else:
                                play_sound("failure")
                                wait_until_green_disappears()
                                print(f'❌第二件不是赠品 (耗时: {time.time() - process_start:.2f}s)')

                    # 产品编号2: Y3 或 Y2
                    elif product_code2 in ['Y3', 'Y2']:
                        # 多线程独立检测第二行（REGION2）是否白色
                        region2_white = is_region_white(REGION2)

                        if region2_white:
                            play_sound("newProduct")
                            click_return_and_wait()
                            print(f'✅新品-无赠品 (耗时: {time.time() - process_start:.2f}s)')
                        else:
                            play_sound("failure")
                            wait_until_green_disappears()
                            print(f'❌新品-有多件 (耗时: {time.time() - process_start:.2f}s)')

                    # 其他产品
                    else:
                        # 多线程独立检测第二行（REGION2）是否白色
                        region2_white = is_region_white(REGION2)

                        if region2_white:
                            play_sound("alternate")
                            click_return_and_wait()
                            print(f'✅其他产品 (耗时: {time.time() - process_start:.2f}s)')
                        else:
                            play_sound("failure")
                            wait_until_green_disappears()
                            print(f'❌其他产品-有多件 (耗时: {time.time() - process_start:.2f}s)')

                else:
                    # 产品数量不是1
                    play_sound("failure")
                    wait_until_green_disappears()
                    print(f'❌产品数量不是1 (耗时: {time.time() - process_start:.2f}s)')

            time.sleep(0.1)

    except Exception as e:
        print(f"监控异常: {e}")
        play_sound("failure")
    finally:
        print("监控循环退出")


def play_sound(sound_type):
    """播放提示音"""
    if sound_type == "success":
        winsound.Beep(2000, 300)
    elif sound_type == "newProduct":
        winsound.Beep(1800, 100)
        winsound.Beep(1800, 100)
    elif sound_type == "alternate":
        winsound.Beep(2500, 200)
    else:
        winsound.Beep(800, 200)
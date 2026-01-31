import numpy as np
import pytesseract
from PIL import Image, ImageOps
import mss
import time
import concurrent.futures
from function_config_manager import load_config, get_current_region

config = load_config()

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Tesseract配置
TESSERACT_CONFIGS = {
    "letter": r'--psm 7 --oem 3 -c tessedit_char_whitelist=ABCD textord_min_linesize=2.5',
    "digit": r'--psm 7 --oem 3 -c tessedit_char_whitelist=123456789',
    "extended": r'--psm 7 --oem 3 -c tessedit_char_whitelist=ABCDJ0123+',
    "NewProduct": r'--psm 7 --oem 3 -c tessedit_char_whitelist=Y23',
}

def preprocess_image(img):
    """优化的图像预处理流水线"""
    # 快速转换为灰度图
    gray = img.convert('L')

    # 使用numpy加速阈值计算
    arr = np.array(gray)
    threshold = max(50, min(np.mean(arr) - 25, 150))

    # 二值化处理
    binary = arr > threshold
    binary_img = Image.fromarray((binary * 255).astype('uint8'))

    # 自动反相判断
    if np.mean(binary) < 0.5:
        binary_img = ImageOps.invert(binary_img)

    # 适度放大图像 (2倍而非3倍)
    return binary_img.resize((img.width * 2, img.height * 2), Image.LANCZOS)

def process_region(config):
    """处理单个区域的优化函数"""
    try:
        # 使用全局mss实例 (线程安全)
        with mss.mss() as sct:
            x, y, w, h = config["position"]
            screenshot = sct.grab({"left": x, "top": y, "width": w, "height": h})

            # 快速图像转换
            img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")

            # 优化的预处理
            processed_img = preprocess_image(img)

            # 执行OCR识别
            text = pytesseract.image_to_string(
                processed_img,
                config=TESSERACT_CONFIGS[config["config_key"]]
            ).strip().replace(' ', '')

            # 根据配置键动态设置白名单
            if config["config_key"] == "letter":
                whitelist = "ABCD"
            elif config["config_key"] == "digit":
                whitelist = "123456789"
            elif config["config_key"] == "extended":
                whitelist = "ABCDJ0123+"
            elif config["config_key"] == "NewProduct":
                whitelist = "Y23"
            else:
                whitelist = ""

            clean_text = ''.join(c for c in text if c in whitelist)

            return {
                "config_key": config["config_key"],  # 新增
                "region": config["name"],
                "result": clean_text or "无有效内容"
            }

    except Exception as e:
        return {
            "config_key": config["config_key"] if config else "unknown",  # 新增
            "region": config["name"] if config else "unknown",
            "result": f"识别异常: {str(e)}"
        }

def get_region_configs():
    """动态获取区域配置（整合原有区域和新区域）"""
    x1, x2 = get_current_region()  # 实时获取最新坐标

    return [
        # 产品编码（ABCD）
        {
            "name": "区域1(ABCD)",
            "position": (x1 + 215, 202, 50, 30),
            "config_key": "letter"
        },
        # 产品数量（123456789）
        {
            "name": "区域2(123456789)",
            "position": (x2 + 1, 194, 20, 30),
            "config_key": "digit"
        },
        # 区域（ABCDJ0123+）
        {
            "name": "区域3(ABCDJ0123+)",
            "position": (x1 + 35, 310, 80, 30),
            "config_key": "extended"
        },
        # 新品区域-4
        {
            "name": "区域5(Y23)",
            "position": (x1 + 5, 202, 20, 30),
            "config_key": "NewProduct"
        },

    ]


def get_region_config_by_key(config_key: str):
    """根据配置键获取单个区域的配置"""
    all_configs = get_region_configs()
    for config in all_configs:
        if config["config_key"] == config_key:
            return config
    return None


def ocr_single_region(config_key: str):
    """识别单个指定区域"""
    config = get_region_config_by_key(config_key)
    if not config:
        return {f"错误": f"未找到配置键 {config_key}"}

    start_time = time.perf_counter()
    result = process_region(config)
    elapsed_time = (time.perf_counter() - start_time) * 1000

    # 返回统一格式的结果字典
    return {
        "config_key": config_key,
        "result": result["result"],
        "time": elapsed_time
    }


def ocr_parallel_scan(region_keys=None):
    """支持选择性区域识别"""
    start_time = time.perf_counter()

    if region_keys is None:
        # 默认扫描所有区域
        current_configs = get_region_configs()
    else:
        # 按需扫描指定区域
        current_configs = [c for c in get_region_configs() if c["config_key"] in region_keys]

    with concurrent.futures.ThreadPoolExecutor() as executor:
        results = list(executor.map(process_region, current_configs))

    elapsed_time = (time.perf_counter() - start_time) * 1000

    # 转换为 {config_key: result} 的字典格式
    results_dict = {r["config_key"]: r["result"] for r in results}  # 现在可以正确访问config_key
    return results_dict, elapsed_time


if __name__ == "__main__":
    # 预热运行
    for _ in range(2):
        ocr_parallel_scan()

    # 正式执行
    scan_results, time_consumed = ocr_parallel_scan()

    print("OCR识别结果：")
    for config_key, result in scan_results.items():  # 修改为遍历字典
        print(f"{config_key}: {result}")

    print(f"\n总识别耗时: {time_consumed:.2f} 毫秒")
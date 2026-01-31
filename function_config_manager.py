import json
from pathlib import Path
from typing import Tuple, TypedDict, Any, Dict
from threading import Lock


class ConfigDict(TypedDict):
    region1_x: int
    region2_x: int
    switch1_state: bool
    switch2_state: bool
    switch3_state: bool
    window_position: Tuple[int, int]
    today_count: int
    last_count_day: int  # 仅用于每日重置


CONFIG_PATH = Path("D:/data/config.json")
DEFAULT_CONFIG: ConfigDict = {
    "region1_x": 185,
    "region2_x": 1676,
    "switch1_state": False,
    "switch2_state": False,
    "switch3_state": True,
    "window_position": [200, 200],
    "today_count": 0,
    "last_count_day": 0  # 仅记录日期
}

# 线程安全的全局状态
_CURRENT_REGION: Tuple[int, int] = (DEFAULT_CONFIG["region1_x"], DEFAULT_CONFIG["region2_x"])
_CONFIG_LOCK = Lock()


def get_config_path() -> Path:
    """获取配置文件路径，保存在D:/data/config.json"""
    return CONFIG_PATH


def load_config() -> ConfigDict:
    """加载或创建默认配置"""
    if not CONFIG_PATH.exists():
        save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG.copy()

    with CONFIG_PATH.open('r', encoding='utf-8') as f:
        config: Dict[str, Any] = json.load(f)
        # 验证配置完整性
        for key in DEFAULT_CONFIG:
            if key not in config:
                config[key] = DEFAULT_CONFIG[key]
        return config  # type: ignore


def save_config(config: ConfigDict) -> None:
    """保存配置到文件"""
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with CONFIG_PATH.open('w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)


def get_current_region() -> Tuple[int, int]:
    """获取当前坐标（线程安全）"""
    with _CONFIG_LOCK:
        return _CURRENT_REGION


def update_ocr_config(x1: int, x2: int) -> None:
    """
    更新坐标并立即生效

    Args:
        x1: 区域1的x坐标
        x2: 区域2的x坐标
    """
    global _CURRENT_REGION

    with _CONFIG_LOCK:
        _CURRENT_REGION = (x1, x2)
        # 保存到配置文件
        config = load_config()
        config.update({"region1_x": x1, "region2_x": x2})
        save_config(config)

# 初始化配置
_init_config = load_config()
_CURRENT_REGION = (_init_config["region1_x"], _init_config["region2_x"])
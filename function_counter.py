import time
import threading
from datetime import datetime, timedelta
from function_config_manager import load_config, save_config


class CounterManager:
    _instance = None
    _lock = threading.Lock()

    def __new__(cls):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super().__new__(cls)
                cls._instance._init_counters()
            return cls._instance

    def _init_counters(self):
        """初始化计数器（程序启动时调用）"""
        config = load_config()
        now = datetime.now()

        # 内存状态初始化（强制重置临时数据）
        self.session_count = 0           # 本次录入量（内存存储）
        self._last_click_time = 0        # 最后操作时间戳（内存存储）
        self._speed_data = []            # 速度计算数据（内存存储）
        self._last_valid_time = 0        # 有效录入时间（内存存储）

        # 每日5点重置检查（仅影响持久化的今日总量）
        current_day = self._get_current_day()
        if config["last_count_day"] != current_day:
            config["today_count"] = 0
            config["last_count_day"] = current_day
            save_config(config)

    def _get_current_day(self):
        """获取当前日期（凌晨5点为分界点）"""
        now = datetime.now()
        if now.hour < 5:
            return (now - timedelta(days=1)).day
        return now.day

    def increment(self):
        """标准计数方法（立即生效）"""
        with self._lock:
            self._unsafe_increment()

    def increment_debounced(self):
        """
        防抖计数方法（1秒内只生效一次）
        返回bool：True表示本次计数有效，False表示被忽略
        """
        with self._lock:
            current_time = time.time()
            if current_time - self._last_click_time < 1:
                return False
            self._unsafe_increment()
            return True

    def _unsafe_increment(self):
        """实际执行计数操作的内部方法（需在锁内调用）"""
        current_time = time.time()
        config = load_config()

        # 检查会话超时（3分钟无操作则重置会话计数）
        if current_time - self._last_click_time > 180:
            self.session_count = 0
            self._last_valid_time = 0

        # 每日5点重置检查
        current_day = self._get_current_day()
        if config["last_count_day"] != current_day:
            config["today_count"] = 0
            config["last_count_day"] = current_day
            save_config(config)

        # 更新计数
        config["today_count"] += 1
        self.session_count += 1          # 仅更新内存中的本次录入量
        self._last_click_time = current_time  # 更新内存中的时间戳
        save_config(config)              # 仅保存今日总量和日期

        # 记录速度数据（连续录入间隔≤7秒时）
        if self._last_valid_time == 0 or (current_time - self._last_valid_time) <= 7:
            self._speed_data.append((current_time, self.session_count))
        self._last_valid_time = current_time

        # 清理超过5分钟的速度数据
        self._clean_speed_data(current_time)

    def _clean_speed_data(self, current_time):
        """清理过期速度数据"""
        cutoff = current_time - 300  # 5分钟前的数据
        self._speed_data = [(t, c) for t, c in self._speed_data if t >= cutoff]
        if not self._speed_data:
            self._speed_data.append((current_time, self.session_count))

    def calculate_speed(self):
        """计算平均录入速度（秒/单）"""
        if len(self._speed_data) < 2:
            return None

        # 分段计算连续录入速度
        valid_segments = []
        current_segment = []
        sorted_data = sorted(self._speed_data, key=lambda x: x[0])

        for point in sorted_data:
            if current_segment and (point[0] - current_segment[-1][0]) > 7:
                if len(current_segment) >= 2:
                    valid_segments.append(current_segment)
                current_segment = []
            current_segment.append(point)

        if len(current_segment) >= 2:
            valid_segments.append(current_segment)

        if not valid_segments:
            return None

        # 计算加权平均速度
        total_time = sum(seg[-1][0] - seg[0][0] for seg in valid_segments)
        total_count = sum(len(seg) - 1 for seg in valid_segments)
        return total_time / total_count if total_count > 0 else None

    def get_counts(self):
        """获取当前计数状态"""
        with self._lock:
            config = load_config()
            current_time = time.time()

            # 运行时检查会话超时（仅重置会话计数）
            if current_time - self._last_click_time > 180:
                self.session_count = 0
                self._last_valid_time = 0

            # 返回今日总量（持久化）和本次录入量（内存）
            return config["today_count"], self.session_count


# 全局单例访问点
counter_manager = CounterManager()
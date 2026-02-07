# sub_table_handler.py
import pandas as pd
from collections import defaultdict
import os


class SubTableHandler:
    def __init__(self, desktop_path):
        self.desktop_path = desktop_path
        self.excluded_waves = set()
        self.brush_waves = set()
        self.sub_table_messages = []
        self.processed_excluded_waves = set()
        self.processed_brush_waves = set()
        self.COLOR_WARN = "purple"
        self.COLOR_ERROR = "red"

    def is_empty(self, value):
        if pd.isna(value):
            return True
        if isinstance(value, str) and value.strip() == "":
            return True
        return False

    def load_sub_table(self):
        sub_table_path = os.path.join(self.desktop_path, "副表.xlsx")
        sub_table_exists = True

        if os.path.exists(sub_table_path):
            try:
                df_sub = pd.read_excel(sub_table_path, sheet_name="副表数据")
                duplicate_messages = []

                if '排除' in df_sub.columns and '刷单' in df_sub.columns:
                    common_values = set(df_sub['排除'].dropna().astype(str).str.strip()) & \
                                    set(df_sub['刷单'].dropna().astype(str).str.strip())
                    if common_values:
                        duplicate_messages.append(f"排除和刷单列存在重复数据: {', '.join(common_values)}")

                if '排除' in df_sub.columns:
                    exclude_duplicates = df_sub['排除'].duplicated()
                    if exclude_duplicates.any():
                        dup_values = df_sub.loc[exclude_duplicates, '排除'].dropna().unique()
                        if len(dup_values) > 0:
                            duplicate_messages.append(f"排除列中存在重复数据: {', '.join(map(str, dup_values))}")

                if '刷单' in df_sub.columns:
                    brush_duplicates = df_sub['刷单'].duplicated()
                    if brush_duplicates.any():
                        dup_values = df_sub.loc[brush_duplicates, '刷单'].dropna().unique()
                        if len(dup_values) > 0:
                            duplicate_messages.append(f"刷单列中存在重复数据: {', '.join(map(str, dup_values))}")

                if duplicate_messages:
                    return "副表数据错误:\n" + "\n".join(duplicate_messages)

                if '排除' in df_sub.columns:
                    self.excluded_waves = set(df_sub['排除'].dropna().astype(str).str.strip().unique())
                if '刷单' in df_sub.columns:
                    self.brush_waves = set(df_sub['刷单'].dropna().astype(str).str.strip().unique())
            except Exception as e:
                self.sub_table_messages.append({"text": f"注意：读取副表时出错: {str(e)}", "color": self.COLOR_WARN})
                sub_table_exists = False
        else:
            self.sub_table_messages.append({"text": "注意：副表不存在", "color": self.COLOR_WARN})
            sub_table_exists = False

        if sub_table_exists and not (self.excluded_waves or self.brush_waves):
            self.sub_table_messages.append({"text": "注意：副表无有效波次数据", "color": self.COLOR_WARN})

        return None

    def is_excluded_wave(self, wave_str):
        if wave_str in self.excluded_waves:
            self.processed_excluded_waves.add(wave_str)
            return True
        return False

    def is_brush_wave(self, wave_str):
        if wave_str in self.brush_waves:
            self.processed_brush_waves.add(wave_str)
            return True
        return False

    def get_warnings(self):
        warnings = []

        if self.brush_waves:
            unmatched_brush = self.brush_waves - self.processed_brush_waves
            if unmatched_brush:
                warnings.append({
                    "text": f"警告：副表中存在没有匹配到的刷单波次: {', '.join(sorted(unmatched_brush))}",
                    "color": self.COLOR_ERROR
                })

        if self.excluded_waves:
            unmatched_excluded = self.excluded_waves - self.processed_excluded_waves
            if unmatched_excluded:
                warnings.append({
                    "text": f"警告：副表中存在没有匹配到的排除波次:{'，'.join(sorted(unmatched_excluded))}",
                    "color": self.COLOR_ERROR
                })

        return warnings
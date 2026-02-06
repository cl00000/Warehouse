import pandas as pd
from collections import defaultdict
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill


class ExcelContrastProcessor:
    def __init__(self):
        self.desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        self.code_dict = {}
        self.priority_dict = {}
        self.sort_dict = {}  # New: Dictionary for sort values
        self.extra_shops = set()
        self.excluded_waves = set()
        self.brush_waves = set()
        self.messages = []
        self.sub_table_messages = []  # New: Separate list for sub-table messages
        self.today_str = datetime.now().strftime("%y%m%d")
        self.null_rows = []
        self.has_null_values = False
        self.has_unmatched_codes = False
        self.non_today_waves = set()
        self.non_today_wave_details = defaultdict(list)
        self.excluded_type_waves = defaultdict(lambda: defaultdict(list))
        self.invalid_orders = set()
        self.brush_style = "洗脸巾/其它包数"
        self.filtered_count = 0
        self.filtered_rows = []
        self.processed_excluded_waves = set()
        self.processed_brush_waves = set()
        self.COLOR_INFO = "green"
        self.COLOR_WARN = "purple"
        self.COLOR_ERROR = "red"
        self.unmatched_waves = defaultdict(list)
        self.partial_unmatched_waves = defaultdict(list)
        self.no_channel_waves = defaultdict(list)  # 新增：存储没有渠道标签的波次

    def is_empty(self, value):
        if pd.isna(value):
            return True
        if isinstance(value, str) and value.strip() == "":
            return True
        return False

    def load_data(self):
        file1_path = os.path.join(self.desktop_path, "1.xlsx")
        code_file_path = os.path.join(self.desktop_path, "编码对应关系.xlsx")
        sub_table_path = os.path.join(self.desktop_path, "副表.xlsx")

        try:
            df1 = pd.read_excel(file1_path, sheet_name="Sheet1")
            required_columns = ['打印波次', '店铺', '货品商家编码', '订单类型', '打单员']
            missing_columns = [col for col in required_columns if col not in df1.columns]
            if missing_columns:
                return None, None, f"1.xlsx中缺少必要的列: {', '.join(missing_columns)}"

            key_columns = ['打印波次', '店铺', '货品商家编码', '订单类型', '打单员']
            null_rows = df1[df1[key_columns].isnull().any(axis=1)]
            if not null_rows.empty:
                null_row_numbers = [idx + 2 for idx in null_rows.index]
                return None, None, f"发现空值行，行号: {', '.join(map(str, null_row_numbers))}"

            if '打单员' in df1.columns:
                df1 = df1[df1['打单员'].astype(str).str.contains('打单')]

            if '订单类型' in df1.columns:
                # 定义有效的订单类型（会被处理的类型）
                valid_types = ['网店销售', '订单补发', '线下零售']

                # 找到不在有效类型中的订单类型
                other_types = df1[~df1['订单类型'].astype(str).str.contains('|'.join(valid_types))]
                if not other_types.empty:
                    unique_types = other_types['订单类型'].unique()
                    # 过滤掉空值或空白字符串
                    unique_types = [str(t).strip() for t in unique_types if str(t).strip()]
                    if unique_types:
                        self.messages.append({
                            "text": f"注意：存在其他订单类型: {', '.join(unique_types)}",
                            "color": self.COLOR_ERROR
                        })
                        print(f"发现其他订单类型: {', '.join(unique_types)}")  # 调试输出

            current_mmdd = datetime.now().strftime("%m%d")
            for idx, row in df1.iterrows():
                wave_value = row['打印波次']
                if not self.is_empty(wave_value):
                    wave_str = str(wave_value).strip()
                    if wave_str.startswith("PB25") and len(wave_str) >= 10:
                        wave_mmdd = wave_str[4:8]
                        if wave_mmdd != current_mmdd:
                            self.non_today_waves.add(wave_str)
                            self.non_today_wave_details[wave_str].append(idx + 2)

            if self.non_today_waves:
                self.messages.append({
                    "text": f"发现非当天波次，已忽略: {', '.join(sorted(self.non_today_waves))}",
                    "color": self.COLOR_INFO
                })

            df_code = pd.read_excel(code_file_path, sheet_name="Sheet1")
            df_code["优先级"] = df_code["优先级"].ffill()

            brush_rows = df_code[df_code["货品商家编码"].astype(str).str.strip() == "刷单"]
            if not brush_rows.empty:
                self.brush_style = brush_rows.iloc[0]["名称"]

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
                        return None, None, "副表数据错误:\n" + "\n".join(duplicate_messages)

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

            return df1, df_code, None

        except Exception as e:
            return None, None, f"读取数据时出错: {str(e)}"

    def build_code_mappings(self, df_code):
        self.style_sort = {}  # New: Mapping from style name to sort value
        for _, row in df_code.iterrows():
            code = str(row["货品商家编码"]).strip()
            name = str(row["名称"]).strip()
            pri = int(row["优先级"]) if not pd.isna(row["优先级"]) else 999
            sort_val = int(row["排序"]) if "排序" in df_code.columns and not pd.isna(row.get("排序")) else 999
            self.code_dict[code] = name
            self.priority_dict[code] = pri
            if name not in self.style_sort or sort_val < self.style_sort[name]:
                self.style_sort[name] = sort_val  # Take the minimum sort value if duplicates

    def get_order_channel(self, shop, wave_str=None, row_idx=None):
        """根据店铺获取订单渠道"""
        shop = str(shop).strip() if not pd.isna(shop) else ""

        # 检查是否有/号
        if '/' not in shop:
            # 如果没有/号，直接返回"其他"，不再记录波次信息
            return "其他"

        # 取第一个/号前的内容
        channel_part = shop.split('/')[0]

        # 新增：如果店铺名第一个/号前两个字为"补发"，则视为"自营"
        if channel_part.startswith("补发"):
            return "自营"

        # 取前两个字作为订单渠道
        if len(channel_part) >= 2:
            return channel_part[:2]
        else:
            return channel_part

    def get_order_type(self, tag):
        """根据订单类型标签获取订单类型"""
        tag = str(tag).strip() if not pd.isna(tag) else ""

        if "订单补发" in tag:
            return "补发单"
        elif "网店销售" in tag or "线下零售" in tag:
            return "新订单"
        else:
            return None

    def parse_codes_and_quantities(self, product_code):
        codes = []
        quantities = {}
        parts = [p.strip() for p in str(product_code).split(';') if p.strip()]

        for part in parts:
            if '*' in part:
                code_part, _, qty_part = part.partition('*')
                try:
                    qty = int(qty_part.strip())
                    code = code_part.strip()
                    codes.append(code)
                    quantities[code] = qty
                except:
                    codes.append(part)
                    quantities[part] = 1
            else:
                codes.append(part)
                quantities[part] = 1

        filtered_codes = [code for code in codes if code not in self.priority_dict or self.priority_dict[code] != 100]
        return filtered_codes, quantities

    class UniformSelector:
        def __init__(self):
            self.group_state = defaultdict(lambda: {
                'index': 0,
                'sorted_codes': None,
                'counters': defaultdict(int),
                'max_qty_codes': None
            })

        def select(self, candidates, quantities, priority_dict, group_key):
            if not candidates:
                return None

            state = self.group_state[group_key]
            priorities = {priority_dict[code] for code in candidates}
            same_priority = len(priorities) == 1
            qty_values = {quantities[code] for code in candidates}
            same_quantity = len(qty_values) == 1

            if same_priority and same_quantity:
                if state['sorted_codes'] is None:
                    state['sorted_codes'] = sorted(candidates)
                selected_code = state['sorted_codes'][state['index'] % len(state['sorted_codes'])]
                state['index'] += 1
                state['counters'][selected_code] += 1
                return selected_code

            elif same_priority:
                if state['max_qty_codes'] is None:
                    max_qty = max(quantities[code] for code in candidates)
                    state['max_qty_codes'] = [code for code in candidates if quantities[code] == max_qty]

                if len(state['max_qty_codes']) > 1:
                    selected_code = state['max_qty_codes'][state['counters']['_index'] % len(state['max_qty_codes'])]
                    state['counters']['_index'] += 1
                else:
                    selected_code = state['max_qty_codes'][0]
                state['counters'][selected_code] += 1
                return selected_code

            else:
                min_priority = min(priority_dict[code] for code in candidates)
                priority_candidates = [code for code in candidates if priority_dict[code] == min_priority]
                return self.select(priority_candidates, quantities, priority_dict, group_key)

    def get_style(self, product_code, selector):
        product_code = str(product_code).strip()
        if not product_code or product_code == "NaN":
            return "（未匹配到编码）", 0

        code_list, quantities = self.parse_codes_and_quantities(product_code)
        if not code_list:
            self.filtered_count += 1
            return None, 0
        unmatched_codes = [code for code in quantities if code not in self.priority_dict]

        if unmatched_codes:
            self.has_unmatched_codes = True
            if len(unmatched_codes) == len(code_list):
                return "（未匹配到编码）", 0
            else:
                return "（未匹配到部分编码）", 0

        valid_codes = [code for code in code_list if code in self.priority_dict]
        if not valid_codes:
            self.has_unmatched_codes = True
            return "（未匹配到编码）", 0

        sorted_codes = tuple(sorted(valid_codes))
        selected_code = selector.select(valid_codes, quantities, self.priority_dict, sorted_codes)
        if selected_code:
            return self.code_dict[selected_code], quantities[selected_code]
        else:
            return "（未匹配到编码）", 0

    def process_data(self, df1):
        result_data = []
        selector = self.UniformSelector()
        data_list = []
        excluded_row_count = 0
        excluded_waves_with_counts = defaultdict(int)
        wave_stats = {
            'brush_waves': defaultdict(list),
            'priority100_filtered': defaultdict(list),
            'priority100_skipped': defaultdict(list)
        }

        max_sort = max(self.style_sort.values()) if self.style_sort else 999

        for idx, row in df1.iterrows():
            wave_value = row['打印波次']
            wave_str = str(wave_value).strip() if not self.is_empty(wave_value) else None
            shop = str(row.get("店铺", "")).strip()

            if wave_str and wave_str in self.non_today_waves:
                continue

            # 获取订单渠道 - 传入波次和行索引
            order_channel = self.get_order_channel(shop, wave_str, idx)
            # 获取订单类型
            order_type = self.get_order_type(row.get("订单类型", ""))
            if order_type is None:
                continue

            if wave_str and wave_str in self.excluded_waves:
                self.processed_excluded_waves.add(wave_str)
                excluded_row_count += 1
                excluded_waves_with_counts[wave_str] += 1
                continue

            if wave_str and wave_str in self.brush_waves:
                self.processed_brush_waves.add(wave_str)
                wave_stats['brush_waves'][wave_str].append(idx)
                style = self.brush_style
                qty = 1  # Assuming qty=1 for brush waves
                order_type = "新订单"
                data_list.append({
                    "订单渠道": order_channel,
                    "订单类型": order_type,
                    "款式": style,
                    "原始编码": str(row["货品商家编码"]).strip(),
                    "波次": wave_str,
                    "数量": qty
                })
            else:
                code_list, quantities = self.parse_codes_and_quantities(row["货品商家编码"])
                if not code_list and any(code in self.priority_dict for code in quantities):
                    wave_stats['priority100_skipped'][wave_str].append(idx)
                    continue
                elif code_list and len(code_list) < len(quantities):
                    wave_stats['priority100_filtered'][wave_str].append(idx)

                style, qty = self.get_style(row["货品商家编码"], selector)
                if style is None:
                    continue

                if style == "（未匹配到编码）":
                    self.unmatched_waves[wave_str].append(idx + 2)
                elif style == "（未匹配到部分编码）":
                    self.partial_unmatched_waves[wave_str].append(idx + 2)

                data_list.append({
                    "订单渠道": order_channel,
                    "订单类型": order_type,
                    "款式": style,
                    "原始编码": str(row["货品商家编码"]).strip(),
                    "波次": wave_str,
                    "数量": qty
                })

        if self.unmatched_waves:
            unmatched_waves = sorted(self.unmatched_waves.keys())
            unmatched_count = sum(len(rows) for rows in self.unmatched_waves.values())
            self.messages.append({
                "text": f"发现 {unmatched_count} 行数据未匹配到任何编码，波次号：{'、'.join(unmatched_waves)}",
                "color": self.COLOR_ERROR
            })

        if self.partial_unmatched_waves:
            partial_waves = sorted(self.partial_unmatched_waves.keys())
            partial_count = sum(len(rows) for rows in self.partial_unmatched_waves.values())
            self.messages.append({
                "text": f"发现 {partial_count} 行数据部分编码未匹配，波次号：{'、'.join(partial_waves)}",
                "color": self.COLOR_ERROR
            })

        if wave_stats['priority100_filtered']:
            filtered_waves = sorted(wave_stats['priority100_filtered'].keys())
            filtered_count = sum(len(rows) for rows in wave_stats['priority100_filtered'].values())
            self.messages.append({
                "text": f"已处理 {filtered_count} 行数据（排除了其中的优先级100编码）:{'，'.join(filtered_waves)}",
                "color": self.COLOR_INFO
            })

        if wave_stats['priority100_skipped']:
            skipped_waves = sorted(wave_stats['priority100_skipped'].keys())
            skipped_count = sum(len(rows) for rows in wave_stats['priority100_skipped'].values())
            self.messages.append({
                "text": f"已跳过 {skipped_count} 行数据（仅包含优先级100编码）:{'，'.join(skipped_waves)}",
                "color": self.COLOR_INFO
            })

        if wave_stats['brush_waves']:
            brush_waves = sorted(wave_stats['brush_waves'].keys())
            brush_count = sum(len(rows) for rows in wave_stats['brush_waves'].values())
            self.messages.append({
                "text": f"已处理 {brush_count} 条刷单波次:{'，'.join(brush_waves)}",
                "color": self.COLOR_INFO
            })

        if self.processed_excluded_waves:
            excluded_waves_sorted = sorted(self.processed_excluded_waves)
            wave_details = [f"{wave}" for wave in excluded_waves_sorted]
            self.messages.append({
                "text": f"已排除 {excluded_row_count} 条数据，来自波次: {', '.join(wave_details)}",
                "color": self.COLOR_INFO
            })

        # Append sub-table messages at the end
        if self.brush_waves:
            unmatched_brush = self.brush_waves - self.processed_brush_waves
            if unmatched_brush:
                self.sub_table_messages.append({
                    "text": f"警告：副表中存在没有匹配到的刷单波次: {', '.join(sorted(unmatched_brush))}",
                    "color": self.COLOR_ERROR
                })

        if self.excluded_waves:
            unmatched_excluded = self.excluded_waves - self.processed_excluded_waves
            if unmatched_excluded:
                self.sub_table_messages.append({
                    "text": f"警告：副表中存在没有匹配到的排除波次:{'，'.join(sorted(unmatched_excluded))}",
                    "color": self.COLOR_ERROR
                })

        # Add sub-table messages to the main messages list
        self.messages.extend(self.sub_table_messages)

        # 订单渠道优先级
        channel_priority = {"自营": 0, "代发": 1, "分销": 2, "其他": 3}  # 添加"其他"优先级为3
        # 订单类型优先级
        order_type_priority = {"新订单": 0, "补发单": 1}

        if data_list:
            temp_df = pd.DataFrame(data_list)
            group_counts = temp_df.groupby(["订单渠道", "订单类型", "款式"]).agg(
                单量=('数量', 'count'),
                实际数量=('数量', 'sum')
            ).reset_index()

            # 创建排序键：订单渠道优先级 + 订单类型优先级 + 款式排序值
            group_counts["订单渠道优先级"] = group_counts["订单渠道"].map(channel_priority)
            group_counts["订单类型优先级"] = group_counts["订单类型"].map(order_type_priority)

            # 为每个款式添加排序值
            group_counts["款式排序"] = group_counts["款式"].map(self.style_sort)

            # 对于没有在编码对应关系中的款式，设置一个较大的排序值
            group_counts["款式排序"] = group_counts["款式排序"].fillna(max_sort + 1)

            # 按照订单渠道优先级、订单类型优先级和款式排序值升序排列
            group_counts = group_counts.sort_values(["订单渠道优先级", "订单类型优先级", "款式排序"])

            # 删除临时列
            group_counts = group_counts.drop(columns=["订单渠道优先级", "订单类型优先级", "款式排序"])

            for _, row in group_counts.iterrows():
                result_data.append([
                    row["订单渠道"],
                    row["订单类型"],
                    row["款式"],
                    row["单量"],
                    row["实际数量"],
                    row["款式"] == "（未匹配到编码）",
                    row["款式"] == "空值"
                ])

        return result_data, []

    def create_output_excel(self, result_data, extra_result_data=None):
        wb = Workbook()
        ws = wb.active
        ws.title = "汇总结果"
        header_font = Font(color="FFFFFF", bold=True)
        header_fill = PatternFill(start_color="9999FF", end_color="9999FF", fill_type="solid")
        headers = ["订单渠道", "订单类型", "款式", "单量", "实际数量"]
        ws.append(headers)

        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill

        result_df = pd.DataFrame(result_data, columns=[
            "订单渠道", "订单类型", "款式", "单量", "实际数量",
            "未匹配标记", "空值行标记"
        ])

        color_map = {
            "green": Font(color="00AA00"),
            "purple": Font(color="800080"),
            "red": Font(color="FF0000")
        }

        # 订单类型颜色设置
        order_type_colors = {
            "新订单": Font(color="46C26F"),  # 绿色
            "补发单": Font(color="F0A800")  # 橙色
        }

        # 订单渠道颜色设置
        channel_colors = {
            "自营": Font(color="EB5050"),  # 红色
            "代发": Font(color="9999FF"),  # 蓝色
            "分销": Font(color="F0A800"),  # 橙色
            "其他": Font(color="0000FF")  # 新增：其他渠道用蓝色字体
        }

        # Restart data writing for alignment
        ws.delete_rows(2, ws.max_row)  # Clear appended data

        max_main_rows = len(result_df) if result_df is not None else 0

        for i in range(max_main_rows):
            actual_row = i + 2
            main_idx = i
            main_row = [
                result_df.iloc[main_idx]["订单渠道"],
                result_df.iloc[main_idx]["订单类型"],
                result_df.iloc[main_idx]["款式"],
                result_df.iloc[main_idx]["单量"],
                result_df.iloc[main_idx]["实际数量"]
            ]

            # Write main data
            for col, val in enumerate(main_row, start=1):
                ws.cell(row=actual_row, column=col, value=val)

            # Apply colors for unmatched/empty - font red for col 3-5
            if result_df.iloc[main_idx]["空值行标记"]:
                for col in range(3, 6):
                    ws.cell(row=actual_row, column=col).font = color_map["red"]
            elif result_df.iloc[main_idx]["未匹配标记"] or result_df.iloc[main_idx]["款式"] == "（未匹配到部分编码）":
                for col in range(3, 6):
                    ws.cell(row=actual_row, column=col).font = color_map["red"]

            # Apply order channel color to col 1
            cell_a = ws.cell(row=actual_row, column=1)
            val_a = cell_a.value
            if val_a in channel_colors:
                cell_a.font = channel_colors[val_a]

            # Apply order type color to col 2
            cell_b = ws.cell(row=actual_row, column=2)
            val_b = cell_b.value
            if val_b in order_type_colors:
                cell_b.font = order_type_colors[val_b]

        # Now append empty row
        ws.append([])

        if self.messages:
            for msg in self.messages:
                if isinstance(msg, dict):
                    ws.append([msg["text"]])
                    last_row = ws.max_row
                    cell = ws.cell(row=last_row, column=1)
                    cell.font = color_map.get(msg["color"], Font())
                else:
                    ws.append([msg])

        # 设置列宽
        column_widths = {'A': 15, 'B': 15, 'C': 30, 'D': 10, 'E': 15}
        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width

        output_path = os.path.join(self.desktop_path, "输出结果.xlsx")
        wb.save(output_path)
        return output_path

    def open_file_windows(self, file_path):
        try:
            os.startfile(file_path)
            return True
        except Exception as e:
            print(f"自动打开文件失败: {str(e)}")
            return False

    def process(self):
        try:
            df1, df_code, error = self.load_data()
            if error:
                return False, error
            self.build_code_mappings(df_code)
            result_data, extra_result_data = self.process_data(df1)
            output_path = self.create_output_excel(result_data, extra_result_data)
            self.open_file_windows(output_path)
            return True, "处理完成，已自动打开输出文件"
        except Exception as e:
            return False, f"处理过程中出错: {str(e)}"


def group_calculation():
    processor = ExcelContrastProcessor()
    return processor.process()
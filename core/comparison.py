"""
比較和差異顯示功能 - 確保 TABLE 一定顯示
"""
import os
import csv
import gzip
import json
import time
from datetime import datetime
from wcwidth import wcwidth
import config.settings as settings
from utils.logging import _get_display_width
from utils.helpers import get_file_mtime
from core.excel_parser import pretty_formula, extract_external_refs, get_excel_last_author
from core.baseline import load_baseline, baseline_file_path
import logging

def print_aligned_console_diff(old_data, new_data, file_info=None, max_display_changes=0):
    """
    三欄式顯示，能處理中英文對齊，並正確顯示 formula。
    Address 欄固定闊度，Baseline/Current 平均分配。
    """
    try:
        term_width = os.get_terminal_size().columns
    except OSError:
        term_width = 120

    address_col_width = 12
    separators_width = 4
    remaining_width = term_width - address_col_width - separators_width
    baseline_col_width = remaining_width // 2
    current_col_width = remaining_width - baseline_col_width

    def wrap_text(text, width):
        lines = []
        current_line = ""
        current_width = 0
        for char in str(text):
            char_width = wcwidth(char)
            if char_width < 0:
                continue
            if current_width + char_width > width:
                lines.append(current_line)
                current_line = char
                current_width = char_width
            else:
                current_line += char
                current_width += char_width
        if current_line:
            lines.append(current_line)
        return lines or ['']

    def pad_line(line, width):
        line_width = _get_display_width(line)
        if line_width is None:
            line_width = len(str(line))
        padding = width - line_width
        return str(line) + ' ' * padding if padding > 0 else str(line)

    def format_cell(cell_value):
        if cell_value is None or cell_value == {}:
            return "(Empty)"
        if isinstance(cell_value, dict):
            formula = cell_value.get("formula")
            if formula:
                return f"={formula}"
            if "value" in cell_value:
                return repr(cell_value["value"])
        return repr(cell_value)
    
    print()
    print("=" * term_width)
    if file_info:
        filename = file_info.get('filename', 'Unknown')
        worksheet = file_info.get('worksheet', '')
        caption = f"{filename} [Worksheet: {worksheet}]" if worksheet else filename
        for cap_line in wrap_text(caption, term_width):
            print(cap_line)
    print("=" * term_width)

    # [修復 2] 提取作者資訊
    baseline_time = file_info.get('baseline_time', 'N/A')
    current_time = file_info.get('current_time', 'N/A')
    old_author = file_info.get('old_author', 'N/A')
    new_author = file_info.get('new_author', 'N/A')

    header_addr = pad_line("Address", address_col_width)
    # [修復 2] 將作者資訊加入標題
    header_base = pad_line(f"Baseline ({baseline_time} by {old_author})", baseline_col_width)
    header_curr = pad_line(f"Current ({current_time} by {new_author})", current_col_width)
    print(f"{header_addr} | {header_base} | {header_curr}")
    print("-" * term_width)

    all_keys = sorted(list(set(old_data.keys()) | set(new_data.keys())))
    if not all_keys:
        print("(No cell changes)")
    else:
        displayed_changes_count = 0
        for key in all_keys:
            if max_display_changes > 0 and displayed_changes_count >= max_display_changes:
                print(f"...(僅顯示前 {max_display_changes} 個變更，總計 {len(all_keys)} 個變更)...")
                break

            old_val = old_data.get(key)
            new_val = new_data.get(key)

            if old_val is not None and new_val is not None:
                if old_val != new_val:
                    old_text = format_cell(old_val)
                    new_text = "[MOD] " + format_cell(new_val)
                else:
                    # This case should ideally not be displayed if we only show changes,
                    # but keeping it for completeness.
                    old_text = format_cell(old_val)
                    new_text = format_cell(new_val)
            elif old_val is not None:
                old_text = format_cell(old_val)
                new_text = "[DEL] (Deleted)"
            else:
                old_text = "(Empty)"
                new_text = "[ADD] " + format_cell(new_val)

            addr_lines = wrap_text(key, address_col_width)
            old_lines = wrap_text(old_text, baseline_col_width)
            new_lines = wrap_text(new_text, current_col_width)
            num_lines = max(len(addr_lines), len(old_lines), len(new_lines))
            for i in range(num_lines):
                a_line = addr_lines[i] if i < len(addr_lines) else ""
                o_line = old_lines[i] if i < len(old_lines) else ""
                n_line = new_lines[i] if i < len(new_lines) else ""
                formatted_a = pad_line(a_line, address_col_width)
                formatted_o = pad_line(o_line, baseline_col_width)
                formatted_n = n_line
                print(f"{formatted_a} | {formatted_o} | {formatted_n}")
            displayed_changes_count += 1
    print("=" * term_width)
    print()

def format_timestamp_for_display(timestamp_str):
    """
    格式化時間戳為顯示格式：2025-07-12 18:51:34
    """
    if not timestamp_str or timestamp_str == 'N/A':
        return 'N/A'
    
    try:
        if 'T' in timestamp_str:
            if '.' in timestamp_str:
                timestamp_str = timestamp_str.split('.')[0]
            return timestamp_str.replace('T', ' ')
        return timestamp_str
    except ValueError as e:
        logging.error(f"格式化時間戳失敗: {timestamp_str}, 錯誤: {e}")
        return timestamp_str

def compare_excel_changes(file_path, silent=False, event_number=None, is_polling=False):
    """
    [修復 1 & 2] 修正重複顯示問題並整合作者資訊
    """
    try:
        from core.excel_parser import dump_excel_cells_with_timeout
        
        base_name = os.path.basename(file_path)
        
        old_baseline = load_baseline(base_name)
        if not old_baseline:
            if not silent:
                print(f"❌ 找不到基準線: {base_name}")
            return False
        
        current_data = dump_excel_cells_with_timeout(file_path, show_sheet_detail=False, silent=True)
        if not current_data:
            if not silent:
                print(f"❌ 無法讀取檔案: {base_name}")
            return False
        
        baseline_cells = old_baseline.get('cells', {})
        if baseline_cells == current_data:
            # [修復 1] 內容無變化，直接返回 False，停止輪詢中的重複打印
            return False
        
        any_sheet_has_changes = False
        
        # [修復 2] 提前獲取作者資訊
        old_author = old_baseline.get('last_author', 'N/A')
        try:
            new_author = get_excel_last_author(file_path)
        except Exception:
            new_author = 'Unknown'

        for worksheet_name in set(baseline_cells.keys()) | set(current_data.keys()):
            old_ws = baseline_cells.get(worksheet_name, {})
            new_ws = current_data.get(worksheet_name, {})
            
            all_addresses = set(old_ws.keys()) | set(new_ws.keys())
            
            old_display_data = {}
            new_display_data = {}
            sheet_has_changes = False
            
            for addr in all_addresses:
                old_cell = old_ws.get(addr, {})
                new_cell = new_ws.get(addr, {})
                
                # 比較時，同時比較 formula 和 value
                if old_cell != new_cell:
                    sheet_has_changes = True
                    any_sheet_has_changes = True
                    old_display_data[addr] = old_cell
                    new_display_data[addr] = new_cell
            
            if sheet_has_changes and not silent:
                baseline_timestamp = old_baseline.get('timestamp', 'N/A')
                current_timestamp = get_file_mtime(file_path)
                
                print_aligned_console_diff(
                    old_display_data,
                    new_display_data,
                    {
                        'filename': base_name,
                        'worksheet': worksheet_name,
                        'baseline_time': format_timestamp_for_display(baseline_timestamp),
                        'current_time': format_timestamp_for_display(current_timestamp),
                        'old_author': old_author, # [修復 2] 傳遞作者
                        'new_author': new_author, # [修復 2] 傳遞作者
                    },
                    max_display_changes=settings.MAX_CHANGES_TO_DISPLAY
                )
                
                try:
                    log_changes_to_csv(file_path, worksheet_name, old_display_data, new_display_data, old_baseline)
                except OSError as e:
                    logging.error(f"記錄變更到 CSV 時發生 I/O 錯誤: {e}")
        
        # [修復 1] 只有在確實有變更時才更新基準線
        if any_sheet_has_changes and not silent:
            updated_baseline = {
                "last_author": new_author,
                "content_hash": f"updated_{int(time.time())}",
                "cells": current_data,
                "timestamp": datetime.now().isoformat()
            }
            
            from core.baseline import save_baseline
            if not save_baseline(base_name, updated_baseline):
                print(f"[WARNING] 基準線更新失敗: {base_name}")
        
        return any_sheet_has_changes
        
    except Exception as e:
        if not silent:
            logging.error(f"比較過程出錯: {e}")
        return False

def analyze_meaningful_changes(old_ws, new_ws):
    """
    🧠 分析有意義的變更
    """
    meaningful_changes = []
    
    # 找出所有儲存格
    all_addresses = set(old_ws.keys()) | set(new_ws.keys())
    
    for addr in all_addresses:
        old_cell = old_ws.get(addr, {})
        new_cell = new_ws.get(addr, {})
        
        old_val = old_cell.get('value')
        new_val = new_cell.get('value')
        old_formula = old_cell.get('formula')
        new_formula = new_cell.get('formula')
        
        # 🔥 變更類型分析
        change_type = classify_change_type(old_cell, new_cell)
        
        if change_type in ['FORMULA_CHANGE', 'DIRECT_VALUE_CHANGE', 'EXTERNAL_REF_UPDATE', 'CELL_ADDED', 'CELL_DELETED']:
            meaningful_changes.append({
                'address': addr,
                'old_value': old_val,
                'new_value': new_val,
                'old_formula': old_formula,
                'new_formula': new_formula,
                'change_type': change_type
            })
    
    return meaningful_changes

def classify_change_type(old_cell, new_cell):
    """
    🔍 分類變更類型
    """
    old_val = old_cell.get('value')
    new_val = new_cell.get('value')
    old_formula = old_cell.get('formula')
    new_formula = new_cell.get('formula')
    
    # 儲存格新增
    if not old_cell and new_cell:
        return 'CELL_ADDED'
    
    # 儲存格刪除
    if old_cell and not new_cell:
        return 'CELL_DELETED'
    
    # 公式有變更
    if old_formula != new_formula:
        return 'FORMULA_CHANGE'
    
    # 沒有公式，但值有變更（直接輸入的值）
    if not old_formula and not new_formula and old_val != new_val:
        return 'DIRECT_VALUE_CHANGE'
    
    # 有公式，公式沒變，但值有變更
    if old_formula and new_formula and old_formula == new_formula and old_val != new_val:
        # 檢查是否為外部參照
        if has_external_reference(old_formula):
            return 'EXTERNAL_REF_UPDATE'
        else:
            return 'INDIRECT_CHANGE'  # 這類不追蹤
    
    return 'NO_CHANGE'

def has_external_reference(formula):
    """
    🔗 檢查公式是否包含外部參照
    """
    if not formula:
        return False
    
    # 檢查常見的外部參照模式
    external_patterns = [
        r"'\[.*?\]",           # '[檔案名]工作表'!
        r"\[.*?\]",            # [檔案名]工作表!
        r"'.*?\.xlsx?'!",      # '檔案名.xlsx'!
        r"'.*?\.xls?'!",       # '檔案名.xls'!
    ]
    
    import re
    for pattern in external_patterns:
        if re.search(pattern, formula, re.IGNORECASE):
            return True
    
    return False

def print_meaningful_changes(changes, file_info):
    """
    📊 顯示有意義的變更
    """
    try:
        term_width = os.get_terminal_size().columns
    except OSError:
        term_width = 120
    
    print()
    print("=" * term_width)
    
    filename = file_info.get('filename', 'Unknown')
    worksheet = file_info.get('worksheet', '')
    caption = f"{filename} [Worksheet: {worksheet}] - 有意義的變更"
    print(caption)
    
    print("=" * term_width)
    
    baseline_time = file_info.get('baseline_time', 'N/A')
    current_time = file_info.get('current_time', 'N/A')
    
    print(f"Address      | Change Type          | Baseline ({baseline_time})       | Current ({current_time})")
    print("-" * term_width)
    
    # 變更類型的中文說明
    change_type_labels = {
        'FORMULA_CHANGE': '🔧 公式變更',
        'DIRECT_VALUE_CHANGE': '✏️ 直接輸入',
        'EXTERNAL_REF_UPDATE': '🔗 外部參照更新',
        'CELL_ADDED': '➕ 新增儲存格',
        'CELL_DELETED': '➖ 刪除儲存格'
    }
    
    for change in changes:
        addr = change['address']
        change_type = change['change_type']
        old_val = change['old_value']
        new_val = change['new_value']
        old_formula = change['old_formula']
        new_formula = change['new_formula']
        
        type_label = change_type_labels.get(change_type, change_type)
        
        if change_type == 'FORMULA_CHANGE':
            old_display = f"[公式] {old_formula}"
            new_display = f"[公式] {new_formula}"
        elif change_type == 'EXTERNAL_REF_UPDATE':
            old_display = f"[外部] {old_val} ({old_formula})"
            new_display = f"[外部] {new_val} ({new_formula})"
        elif change_type == 'CELL_ADDED':
            old_display = "(Empty)"
            new_display = f"[ADD] {new_formula or new_val}"
        elif change_type == 'CELL_DELETED':
            old_display = f"{old_formula or old_val}"
            new_display = "[DEL] (Deleted)"
        else:
            old_display = str(old_val)
            new_display = str(new_val)
        
        print(f"{addr:<12} | {type_label:<20} | {old_display:<30} | {new_display}")
    
    print("=" * term_width)
    print()

def log_meaningful_changes_to_csv(file_path, worksheet_name, changes, baseline_data):
    """
    📝 記錄有意義的變更到 CSV
    """
    try:
        os.makedirs(os.path.dirname(settings.CSV_LOG_FILE), exist_ok=True)
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        file_exists = os.path.exists(settings.CSV_LOG_FILE)
        
        with gzip.open(settings.CSV_LOG_FILE, 'at', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            
            if not file_exists:
                writer.writerow([
                    'Timestamp', 'Filename', 'Worksheet', 'Cell', 
                    'Change_Type', 'Old_Value', 'New_Value', 'Old_Formula', 'New_Formula', 'Last_Author'
                ])
            
            for change in changes:
                writer.writerow([
                    timestamp,
                    os.path.basename(file_path),
                    worksheet_name,
                    change['address'],
                    change['change_type'],
                    change['old_value'],
                    change['new_value'],
                    change['old_formula'],
                    change['new_formula'],
                    baseline_data.get('last_author', 'Unknown')
                ])
        
        print(f"📝 有意義變更已記錄到 CSV")
        
    except (OSError, csv.Error) as e:
        logging.error(f"記錄有意義變更到 CSV 時發生錯誤: {e}")

def update_baseline_after_meaningful_changes(file_path, base_name, current_data):
    """
    🔄 更新基準線
    """
    try:
        from core.excel_parser import get_excel_last_author
        current_author = get_excel_last_author(file_path)
    except Exception as e:
        logging.error(f"獲取 Excel 最後作者失敗: {e}")
        current_author = 'Unknown'
    
    # 更新基準線
    updated_baseline = {
        "last_author": current_author,
        "content_hash": f"updated_{int(time.time())}",
        "cells": current_data,
        "timestamp": datetime.now().isoformat()
    }
    
    # 保存更新的基準線
    from core.baseline import save_baseline
    if save_baseline(base_name, updated_baseline):
        pass  # 成功更新
    else:
        print(f"[WARNING] 基準線更新失敗: {base_name}")






def log_changes_to_csv(file_path, worksheet_name, old_data, new_data, baseline_data):
    """
    記錄變更到 CSV 檔案
    """
    try:
        os.makedirs(os.path.dirname(settings.CSV_LOG_FILE), exist_ok=True)
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        file_exists = os.path.exists(settings.CSV_LOG_FILE)
        
        with gzip.open(settings.CSV_LOG_FILE, 'at', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            
            if not file_exists:
                writer.writerow([
                    'Timestamp', 'Filename', 'Worksheet', 'Cell', 
                    'Old_Value', 'New_Value', 'Last_Author', 'Change_Type'
                ])
            
            all_addresses = set(old_data.keys()) | set(new_data.keys())
            for addr in all_addresses:
                old_val = old_data.get(addr)
                new_val = new_data.get(addr)
                
                if old_val is None and new_val is not None:
                    change_type = 'ADD'
                elif old_val is not None and new_val is None:
                    change_type = 'DEL'
                else:
                    change_type = 'MOD'
                
                writer.writerow([
                    timestamp,
                    os.path.basename(file_path),
                    worksheet_name,
                    addr,
                    old_val,
                    new_val,
                    baseline_data.get('last_author', 'Unknown'),
                    change_type
                ])
        
        print(f"📝 變更已記錄到 CSV")
        
    except (OSError, csv.Error) as e:
        logging.error(f"記錄變更到 CSV 時發生錯誤: {e}")

# 保留輔助函數
def should_filter_change(change):
    old_f, new_f = change.get('old_formula'), change.get('new_formula')
    old_v, new_v = change.get('old_value'), change.get('new_value')
    
    if (old_f is None) and (new_f is None):
        return old_v == new_v
    else:
        return old_f == new_f

def filter_array_formula_change(change):
    old_f, new_f = change.get('old_formula'), change.get('new_formula')
    return old_f == new_f

def enrich_formula_external_path(change, ref_map):
    c = change.copy()
    c['old_formula'] = pretty_formula(c.get('old_formula'), ref_map)
    c['new_formula'] = pretty_formula(c.get('new_formula'), ref_map)
    return c

def set_current_event_number(event_number):
    compare_excel_changes._current_event_number = event_number

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

def print_aligned_console_diff(old_data, new_data, file_info=None):
    """
    全新版本的三欄式顯示，能完美處理中英文對齊問題。
    Address 欄固定一個較小的闊度，剩餘空間由 Baseline 和 Current 平均分享。
    """
    # 嘗試獲取終端機的闊度，如果失敗則使用一個預設值
    try:
        term_width = os.get_terminal_size().columns
    except OSError:
        term_width = 120 # 預設闊度

    # --- 全新、更智能的欄位闊度計算 ---
    # 為 Address 設定一個合理的固定闊度
    address_col_width = 12
    # 兩個分隔符 ' | ' 共佔 4 個字元位
    separators_width = 4
    # 計算剩餘可用於內容顯示的闊度
    remaining_width = term_width - address_col_width - separators_width
    
    # 將剩餘空間盡量平均分配給 Baseline 和 Current
    baseline_col_width = remaining_width // 2
    # Current 欄位得到剩餘的部分，可以處理單數闊度的情況
    current_col_width = remaining_width - baseline_col_width

    # --- 輔助函數，用於文字換行 ---
    def wrap_text(text, width):
        lines = []
        current_line = ""
        current_width = 0
        for char in str(text):
            # 使用 wcwidth 獲取單個字元的闊度
            char_width = wcwidth(char)
            if char_width < 0: # 忽略控制字元
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
        # 如果輸入是空字串，確保返回一個包含空字串的列表，以佔據一行
        return lines or ['']

    # --- 輔助函數，用於將單行文字填充到指定闊度 ---
    def pad_line(line, width):
        # 計算目前行的實際顯示闊度
        line_width = _get_display_width(line)
        if line_width is None:
            line_width = len(str(line))
        # 計算需要填充的空格數量
        padding = width - line_width
        # 返回填充後的字串
        return str(line) + ' ' * padding if padding > 0 else str(line)

    # ==================== 開始打印輸出 ====================
    
    # 表格上方加空行
    print()
    
    # 🔥 表格最頂部 - 用等號
    print("=" * term_width)
    
    # 打印檔案和工作表標題
    if file_info:
        filename = file_info.get('filename', 'Unknown')
        worksheet = file_info.get('worksheet', '')
        caption = f"{filename} [Worksheet: {worksheet}]" if worksheet else filename
        # 標題也需要支援換行
        for cap_line in wrap_text(caption, term_width):
            print(cap_line)
    
    # 🔥 標題下方 - 用等號
    print("=" * term_width)

    # 打印表頭
    baseline_time = file_info.get('baseline_time', 'N/A')
    current_time = file_info.get('current_time', 'N/A')
    
    header_addr = pad_line("Address", address_col_width)
    header_base = pad_line(f"Baseline ({baseline_time})", baseline_col_width)
    header_curr = pad_line(f"Current ({current_time})", current_col_width)
    print(f"{header_addr} | {header_base} | {header_curr}")
    
    # 🔥 表頭下方 - 用橫線
    print("-" * term_width)

    # 準備數據進行比較
    all_keys = sorted(list(set(old_data.keys()) | set(new_data.keys())))

    if not all_keys:
        print("(No cell changes)")
    else:
        for key in all_keys:
            old_val = old_data.get(key)
            new_val = new_data.get(key)
            
            # 準備顯示的文字
            if old_val is not None and new_val is not None:
                old_text = f"'{old_val}'"
                new_text = f"[MOD] '{new_val}'" if old_val != new_val else f"'{new_val}'"
            elif old_val is not None:
                old_text = f"'{old_val}'"
                new_text = "[DEL] (Deleted)"
            else:
                old_text = "(Empty)"
                new_text = f"[ADD] '{new_val}'"

            # 對三欄的內容分別進行文字換行
            addr_lines = wrap_text(key, address_col_width)
            old_lines = wrap_text(old_text, baseline_col_width)
            new_lines = wrap_text(new_text, current_col_width)

            # 計算需要打印的最大行數
            num_lines = max(len(addr_lines), len(old_lines), len(new_lines))

            # 逐行打印，確保每一行都對齊
            for i in range(num_lines):
                # 從換行後的列表中取出對應行的文字，如果該欄沒有那麼多行，則為空字串
                a_line = addr_lines[i] if i < len(addr_lines) else ""
                o_line = old_lines[i] if i < len(old_lines) else ""
                n_line = new_lines[i] if i < len(new_lines) else ""

                # 對每一行的文字進行填充，使其達到該欄的闊度
                formatted_a = pad_line(a_line, address_col_width)
                formatted_o = pad_line(o_line, baseline_col_width)
                # Current 欄位不需要填充，因為它是最右邊的一欄
                formatted_n = n_line

                print(f"{formatted_a} | {formatted_o} | {formatted_n}")

    # 🔥 表格最底部 - 用等號
    print("=" * term_width)
    
    # 表格下方加空行
    print()

def format_timestamp_for_display(timestamp_str):
    """
    格式化時間戳為顯示格式：2025-07-12 18:51:34
    """
    if not timestamp_str or timestamp_str == 'N/A':
        return 'N/A'
    
    try:
        # 如果是 ISO 格式 (2025-07-12T18:51:34.123456)
        if 'T' in timestamp_str:
            # 移除微秒部分，只保留到秒
            if '.' in timestamp_str:
                timestamp_str = timestamp_str.split('.')[0]
            # 將 T 替換為空格
            return timestamp_str.replace('T', ' ')
        
        # 如果已經是正確格式，直接返回
        return timestamp_str
        
    except Exception:
        return timestamp_str

def compare_excel_changes(file_path, silent=False, event_number=None, is_polling=False):
    """
    🔥 強制顯示 TABLE 的簡化版本
    """
    try:
        from core.excel_parser import dump_excel_cells_with_timeout
        
        base_name = os.path.basename(file_path)
        
        # 載入基準線
        baseline_file = baseline_file_path(base_name)
        old_baseline = load_baseline(baseline_file)
        
        if not old_baseline:
            if not silent:
                print(f"❌ 找不到基準線: {base_name}")
            return False
        
        # 讀取當前檔案內容
        current_data = dump_excel_cells_with_timeout(file_path, show_sheet_detail=False, silent=True)
        if not current_data:
            if not silent:
                print(f"❌ 無法讀取檔案: {base_name}")
            return False
        
        # 比較變更
        has_changes = False
        baseline_cells = old_baseline.get('cells', {})
        
        # 為每個工作表進行比較
        for worksheet_name in set(baseline_cells.keys()) | set(current_data.keys()):
            old_ws = baseline_cells.get(worksheet_name, {})
            new_ws = current_data.get(worksheet_name, {})
            
            # 找出所有儲存格
            all_addresses = set(old_ws.keys()) | set(new_ws.keys())
            
            # 🔥 強制準備顯示數據
            old_display_data = {}
            new_display_data = {}
            
            for addr in all_addresses:
                old_cell = old_ws.get(addr, {})
                new_cell = new_ws.get(addr, {})
                
                old_val = old_cell.get('value')
                new_val = new_cell.get('value')
                old_formula = old_cell.get('formula')
                new_formula = new_cell.get('formula')
                
                # 🔥 只要有任何變更就顯示
                if old_val != new_val or old_formula != new_formula:
                    has_changes = True
                    # 🔥 直接顯示值，不管任何設定
                    old_display_data[addr] = old_val
                    new_display_data[addr] = new_val
            
            # 🔥 如果有變更，強制顯示 TABLE
            if (old_display_data or new_display_data) and not silent:
                # 格式化時間顯示
                baseline_timestamp = old_baseline.get('timestamp', 'N/A')
                current_timestamp = get_file_mtime(file_path)
                
                formatted_baseline_time = format_timestamp_for_display(baseline_timestamp)
                formatted_current_time = format_timestamp_for_display(current_timestamp)
                
                print_aligned_console_diff(
                    old_display_data,
                    new_display_data,
                    {
                        'filename': base_name,
                        'worksheet': worksheet_name,
                        'baseline_time': formatted_baseline_time,
                        'current_time': formatted_current_time
                    }
                )
                
                # 記錄變更到 CSV
                try:
                    log_changes_to_csv(file_path, worksheet_name, old_display_data, new_display_data, old_baseline)
                except Exception:
                    pass
        
        return has_changes
        
    except Exception as e:
        if not silent:
            print(f"❌ 比較過程出錯: {e}")
        return False

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
        
    except Exception:
        pass

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
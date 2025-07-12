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
from utils.logging import _get_display_width, get_logger
from utils.helpers import get_file_mtime
from core.excel_parser import pretty_formula, extract_external_refs, get_excel_last_author
from core.baseline import load_baseline, baseline_file_path

# 獲取日誌器
logger = get_logger(__name__)

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
        
    except (ValueError, TypeError) as e:
        logger.debug(f"時間戳格式轉換時發生錯誤：{e}")
        return timestamp_str
    except Exception as e:
        logger.warning(f"格式化時間戳時發生未預期錯誤：{type(e).__name__}: {e}")
        return timestamp_str

def compare_excel_changes(file_path, silent=False, event_number=None, is_polling=False):
    """
    🔥 強制顯示 TABLE 的簡化版本 - 修正重複顯示問題
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
        
        # 🔥 檢查內容是否真的有變化
        baseline_cells = old_baseline.get('cells', {})
        
        # 快速比較 - 如果結構完全相同，跳過顯示
        if baseline_cells == current_data:
            return False
        
        # 比較變更
        has_changes = False
        changes_found = False
        
        # 為每個工作表進行比較
        for worksheet_name in set(baseline_cells.keys()) | set(current_data.keys()):
            old_ws = baseline_cells.get(worksheet_name, {})
            new_ws = current_data.get(worksheet_name, {})
            
            # 找出所有儲存格
            all_addresses = set(old_ws.keys()) | set(new_ws.keys())
            
            # 🔥 準備顯示數據
            old_display_data = {}
            new_display_data = {}
            
            for addr in all_addresses:
                old_cell = old_ws.get(addr, {})
                new_cell = new_ws.get(addr, {})
                
                old_val = old_cell.get('value')
                new_val = new_cell.get('value')
                old_formula = old_cell.get('formula')
                new_formula = new_cell.get('formula')
                
                # 🔥 原本的邏輯：只要有任何變更就顯示
                if old_val != new_val or old_formula != new_formula:
                    has_changes = True
                    changes_found = True
                    
                    # ⭐ 新增：外部參照特殊處理
                    # 如果公式沒變但值有變，且包含外部參照，仍然要追蹤
                    if (old_formula == new_formula and 
                        old_val != new_val and 
                        has_external_reference(old_formula)):
                        print(f"🔗 外部參照更新: {addr} = {old_formula}")
                    
                    # 🔥 保持原本的顯示邏輯不變
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
                except (OSError, IOError) as e:
                    logger.error(f"寫入CSV日誌時發生I/O錯誤：{e}")
                except PermissionError as e:
                    logger.warning(f"無權限寫入CSV日誌：{e}")
                except Exception as e:
                    logger.error(f"記錄變更到CSV時發生未預期錯誤：{type(e).__name__}: {e}")
        
        # 🔥 重要：如果發現變更，立即更新基準線以避免重複顯示
        if has_changes and not silent:
            # 獲取當前檔案的作者
            try:
                current_author = get_excel_last_author(file_path)
            except:
                current_author = 'Unknown'
            
            # 更新基準線
            updated_baseline = {
                "last_author": current_author,
                "content_hash": f"updated_{int(time.time())}",  # 簡單的雜湊
                "cells": current_data,
                "timestamp": datetime.now().isoformat()
            }
            
            # 保存更新的基準線
            from core.baseline import save_baseline
            if save_baseline(base_name, updated_baseline):
                # 不顯示更新訊息，避免太多輸出
                pass
            else:
                print(f"[WARNING] 基準線更新失敗: {base_name}")
        
        return has_changes
        
    except FileNotFoundError as e:
        logger.error(f"Excel檔案不存在：{file_path} - {e}")
        if not silent:
            print(f"❌ 比較過程出錯: 檔案不存在")
        return False
    except PermissionError as e:
        logger.warning(f"無權限訪問Excel檔案：{file_path} - {e}")
        if not silent:
            print(f"❌ 比較過程出錯: 權限被拒絕")
        return False
    except Exception as e:
        logger.error(f"比較Excel變更時發生未預期錯誤：{file_path} - {type(e).__name__}: {e}")
        if not silent:
            print(f"❌ 比較過程出錯: {e}")
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
        
    except FileNotFoundError as e:
        logger.error(f"CSV日誌檔案目錄不存在：{e}")
    except PermissionError as e:
        logger.warning(f"無權限寫入CSV日誌檔案：{e}")
    except (OSError, IOError) as e:
        logger.error(f"寫入CSV日誌時發生I/O錯誤：{e}")
    except Exception as e:
        logger.error(f"記錄有意義變更到CSV時發生未預期錯誤：{type(e).__name__}: {e}")

def update_baseline_after_meaningful_changes(file_path, base_name, current_data):
    """
    🔄 更新基準線
    """
    try:
        from core.excel_parser import get_excel_last_author
        current_author = get_excel_last_author(file_path)
    except:
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
        
    except FileNotFoundError as e:
        logger.error(f"CSV日誌檔案目錄不存在：{e}")
    except PermissionError as e:
        logger.warning(f"無權限寫入CSV日誌檔案：{e}")
    except (OSError, IOError) as e:
        logger.error(f"寫入CSV日誌時發生I/O錯誤：{e}")
    except Exception as e:
        logger.error(f"記錄變更到CSV時發生未預期錯誤：{type(e).__name__}: {e}")

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
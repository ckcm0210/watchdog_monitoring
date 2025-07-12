"""
通用輔助函數
"""
import os
import time
import json
import threading
from datetime import datetime
import config.settings as settings

def get_file_mtime(filepath):
    """
    獲取檔案修改時間
    """
    try:
        return datetime.fromtimestamp(os.path.getmtime(filepath)).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return "Unknown"

def human_readable_size(num_bytes):
    """
    轉換檔案大小為人類可讀格式
    """
    if num_bytes is None: 
        return "0 B"
    
    for unit in ['B','KB','MB','GB','TB']:
        if num_bytes < 1024.0: 
            return f"{num_bytes:,.2f} {unit}"
        num_bytes /= 1024.0
    return f"{num_bytes:.2f} PB"

def get_all_excel_files(folders):
    """
    獲取所有Excel檔案
    """
    all_files = []
    for folder in folders:
        if os.path.isfile(folder):
            if folder.lower().endswith(settings.SUPPORTED_EXTS) and not os.path.basename(folder).startswith('~$'):
                all_files.append(folder)
        elif os.path.isdir(folder):
            for dirpath, _, filenames in os.walk(folder):
                for f in filenames:
                    if f.lower().endswith(settings.SUPPORTED_EXTS) and not f.startswith('~$'):
                        all_files.append(os.path.join(dirpath, f))
    return all_files

def is_force_baseline_file(filepath):
    """
    檢查是否為強制baseline檔案
    """
    try:
        for pattern in settings.FORCE_BASELINE_ON_FIRST_SEEN:
            if pattern.lower() in filepath.lower(): 
                return True
        return False
    except Exception: 
        return False

def save_progress(completed_files, total_files):
    """
    保存進度
    """
    if not settings.ENABLE_RESUME: 
        return
    
    try:
        progress_data = {
            "timestamp": datetime.now().isoformat(), 
            "completed": completed_files, 
            "total": total_files
        }
        
        # 確保目錄存在
        os.makedirs(os.path.dirname(settings.RESUME_LOG_FILE), exist_ok=True)
        
        with open(settings.RESUME_LOG_FILE, 'w', encoding='utf-8') as f: 
            json.dump(progress_data, f, ensure_ascii=False, indent=2)
    except Exception as e: 
        print(f"[WARN] 無法儲存進度: {e}")

def load_progress():
    """
    載入進度
    """
    if not settings.ENABLE_RESUME or not os.path.exists(settings.RESUME_LOG_FILE): 
        return None
    
    try:
        with open(settings.RESUME_LOG_FILE, 'r', encoding='utf-8') as f: 
            return json.load(f)
    except Exception as e:
        print(f"[WARN] 無法載入進度: {e}")
        return None

def timeout_handler():
    """
    超時處理器
    """
    while not settings.force_stop and not settings.baseline_completed:
        time.sleep(10)
        if settings.current_processing_file and settings.processing_start_time:
            elapsed = time.time() - settings.processing_start_time
            if elapsed > settings.FILE_TIMEOUT_SECONDS:
                print(f"\n⏰ 檔案處理超時! (檔案: {settings.current_processing_file}, 已處理: {elapsed:.1f}s > {settings.FILE_TIMEOUT_SECONDS}s)")
                settings.current_processing_file = None
                settings.processing_start_time = None
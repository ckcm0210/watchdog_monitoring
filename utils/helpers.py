"""
通用輔助函數
"""
import os
import time
import json
import threading
from datetime import datetime
import config.settings as settings
from utils.logging import get_logger

# 獲取日誌器
logger = get_logger(__name__)

def get_file_mtime(filepath):
    """
    獲取檔案修改時間
    """
    try:
        return datetime.fromtimestamp(os.path.getmtime(filepath)).strftime("%Y-%m-%d %H:%M:%S")
    except FileNotFoundError:
        logger.warning(f"檔案不存在，無法獲取修改時間：{filepath}")
        return "Unknown"
    except PermissionError:
        logger.warning(f"無權限訪問檔案，無法獲取修改時間：{filepath}")
        return "Unknown"
    except OSError as e:
        logger.error(f"獲取檔案修改時間時發生系統錯誤：{filepath} - {e}")
        return "Unknown"
    except (ValueError, OverflowError) as e:
        logger.error(f"檔案修改時間無效：{filepath} - {e}")
        return "Unknown"
    except Exception as e:
        logger.error(f"獲取檔案修改時間時發生未預期錯誤：{filepath} - {type(e).__name__}: {e}")
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
    except (AttributeError, TypeError) as e:
        logger.warning(f"檢查強制baseline檔案時發生設定錯誤：{filepath} - {e}")
        return False
    except Exception as e:
        logger.error(f"檢查強制baseline檔案時發生未預期錯誤：{filepath} - {type(e).__name__}: {e}")
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
    except FileNotFoundError as e:
        logger.error(f"無法創建進度檔案目錄：{e}")
        print(f"[WARN] 無法儲存進度: 目錄不存在")
    except PermissionError as e:
        logger.warning(f"無權限寫入進度檔案：{e}")
        print(f"[WARN] 無法儲存進度: 權限被拒絕")
    except (OSError, IOError) as e:
        logger.error(f"儲存進度時發生I/O錯誤：{e}")
        print(f"[WARN] 無法儲存進度: I/O錯誤")
    except (TypeError, ValueError) as e:
        logger.error(f"序列化進度數據時發生錯誤：{e}")
        print(f"[WARN] 無法儲存進度: 數據錯誤")
    except Exception as e:
        logger.error(f"儲存進度時發生未預期錯誤：{type(e).__name__}: {e}")
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
    except FileNotFoundError:
        logger.debug("進度檔案不存在")
        return None
    except PermissionError as e:
        logger.warning(f"無權限讀取進度檔案：{e}")
        print(f"[WARN] 無法載入進度: 權限被拒絕")
        return None
    except (OSError, IOError) as e:
        logger.error(f"讀取進度檔案時發生I/O錯誤：{e}")
        print(f"[WARN] 無法載入進度: I/O錯誤")
        return None
    except json.JSONDecodeError as e:
        logger.error(f"進度檔案格式錯誤：{e}")
        print(f"[WARN] 無法載入進度: 檔案格式錯誤")
        return None
    except Exception as e:
        logger.error(f"載入進度時發生未預期錯誤：{type(e).__name__}: {e}")
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
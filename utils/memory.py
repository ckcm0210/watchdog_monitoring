"""
記憶體監控功能
"""
import psutil
import os
import gc
import config.settings as settings
from utils.logging import get_logger

# 獲取日誌器
logger = get_logger(__name__)

def get_memory_usage():
    """
    獲取當前記憶體使用量 (MB)
    """
    try:
        return psutil.Process(os.getpid()).memory_info().rss / 1024 / 1024
    except psutil.NoSuchProcess:
        logger.error("無法獲取記憶體使用量：進程不存在")
        return 0
    except psutil.AccessDenied:
        logger.warning("無法獲取記憶體使用量：權限被拒絕")
        return 0
    except OSError as e:
        logger.error(f"獲取記憶體使用量時發生系統錯誤：{e}")
        return 0
    except Exception as e:
        logger.error(f"獲取記憶體使用量時發生未預期錯誤：{type(e).__name__}: {e}")
        return 0

def check_memory_limit():
    """
    檢查記憶體使用是否超過限制
    """
    if not settings.ENABLE_MEMORY_MONITOR: 
        return False
    
    current_memory = get_memory_usage()
    if current_memory > settings.MEMORY_LIMIT_MB:
        print(f"⚠️ 記憶體使用量過高: {current_memory:.1f} MB > {settings.MEMORY_LIMIT_MB} MB")
        print("   正在執行垃圾回收...")
        gc.collect()
        new_memory = get_memory_usage()
        print(f"   垃圾回收後: {new_memory:.1f} MB")
        return new_memory > settings.MEMORY_LIMIT_MB
    return False
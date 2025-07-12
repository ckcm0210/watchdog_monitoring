"""
檔案緩存管理
"""
import os
import time
import hashlib
import shutil
import config.settings as settings
from utils.logging import get_logger

# 獲取日誌器
logger = get_logger(__name__)

def copy_to_cache(network_path, silent=False):
    """
    將網路檔案複製到本地緩存
    """
    if not settings.USE_LOCAL_CACHE: 
        return network_path
    
    try:
        os.makedirs(settings.CACHE_FOLDER, exist_ok=True)
        
        if not os.path.exists(network_path): 
            raise FileNotFoundError(f"網絡檔案不存在: {network_path}")
        
        if not os.access(network_path, os.R_OK): 
            raise PermissionError(f"無法讀取網絡檔案: {network_path}")
        
        file_hash = hashlib.md5(network_path.encode('utf-8')).hexdigest()[:16]
        cache_file = os.path.join(settings.CACHE_FOLDER, f"{file_hash}_{os.path.basename(network_path)}")
        
        if os.path.exists(cache_file):
            try:
                if os.path.getmtime(cache_file) >= os.path.getmtime(network_path): 
                    return cache_file
            except OSError as e:
                logger.debug(f"比較檔案修改時間時發生錯誤：{e}")
                pass
            except Exception as e:
                logger.warning(f"檢查緩存檔案時發生未預期錯誤：{type(e).__name__}: {e}")
                pass
        
        network_size = os.path.getsize(network_path)
        if not silent: 
            print(f"   📥 複製到緩存: {os.path.basename(network_path)} ({network_size/(1024*1024):.1f} MB)")
        
        copy_start = time.time()
        shutil.copy2(network_path, cache_file)
        
        if not silent: 
            print(f"      複製完成，耗時 {time.time() - copy_start:.1f} 秒")
        
        return cache_file
        
    except FileNotFoundError:
        logger.error(f"網絡檔案不存在：{network_path}")
        if not silent: 
            print(f"   ❌ 緩存失敗：檔案不存在")
        return network_path
    except PermissionError as e:
        logger.warning(f"無權限訪問檔案：{network_path} - {e}")
        if not silent: 
            print(f"   ❌ 緩存失敗：權限被拒絕")
        return network_path
    except OSError as e:
        logger.error(f"緩存檔案時發生系統錯誤：{e}")
        if not silent: 
            print(f"   ❌ 緩存失敗：系統錯誤 - {e}")
        return network_path
    except shutil.Error as e:
        logger.error(f"複製檔案時發生錯誤：{e}")
        if not silent: 
            print(f"   ❌ 緩存失敗：複製錯誤 - {e}")
        return network_path
    except Exception as e:
        logger.error(f"緩存檔案時發生未預期錯誤：{type(e).__name__}: {e}")
        if not silent: 
            print(f"   ❌ 緩存失敗：{type(e).__name__}: {e}")
        return network_path
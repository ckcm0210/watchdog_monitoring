"""
基準線管理功能
"""
import os
import json
import gzip
import shutil
import time
import gc
import threading
from datetime import datetime
import config.settings as settings
from utils.helpers import save_progress, load_progress
from utils.memory import check_memory_limit, get_memory_usage
from core.excel_parser import dump_excel_cells_with_timeout, hash_excel_content, get_excel_last_author

def baseline_file_path(base_name):
    """
    獲取基準線檔案路徑
    """
    return os.path.join(settings.LOG_FOLDER, f"{base_name}.baseline.json.gz")

def load_baseline(baseline_file):
    """
    載入基準線檔案，確保文件句柄被正確釋放
    """
    try:
        if not os.path.exists(baseline_file):
            return None
            
        # 加入文件鎖檢查
        try:
            with open(baseline_file, 'r+b') as test_file:
                pass  # 只是測試是否可以存取
        except (PermissionError, OSError) as e:
            print(f"[WARN] Baseline 文件被鎖定: {baseline_file} - {e}")
            return None
            
        # 使用 with 語句確保文件被正確關閉
        with gzip.open(baseline_file, 'rt', encoding='utf-8') as f:
            data = json.load(f)
            
        # 強制等待文件系統釋放句柄
        time.sleep(0.1)
        
        return data
        
    except Exception as e:
        print(f"[ERROR][load_baseline] {baseline_file}: {e}")
        return None

def save_baseline(baseline_file, data):
    """
    保存基準線檔案，強化版本確保文件句柄被正確釋放
    """
    dir_name = os.path.dirname(baseline_file)
    os.makedirs(dir_name, exist_ok=True)
    
    max_retries = 5
    base_delay = 0.2
    
    for attempt in range(max_retries):
        temp_file = None
        try:
            # 使用唯一的臨時文件名
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
            temp_file = os.path.join(dir_name, f"baseline_temp_{timestamp}_{attempt}.tmp")
            
            # 加入時間戳記到 baseline 數據中
            data_with_timestamp = data.copy()
            data_with_timestamp['timestamp'] = datetime.now().isoformat()
            
            # 寫入臨時文件，確保文件句柄被釋放
            with gzip.open(temp_file, 'wt', encoding='utf-8') as f:
                json.dump(data_with_timestamp, f, ensure_ascii=False, separators=(',', ':'))
            
            # 強制刷新文件系統
            time.sleep(0.1)
            
            # 驗證臨時文件完整性
            with gzip.open(temp_file, 'rt', encoding='utf-8') as f:
                json.load(f)
            
            # 強制等待文件系統釋放句柄
            time.sleep(0.1)
            
            # 如果目標文件存在，先備份再刪除
            backup_file = None
            if os.path.exists(baseline_file):
                backup_file = f"{baseline_file}.backup_{timestamp}"
                try:
                    shutil.copy2(baseline_file, backup_file)
                    os.remove(baseline_file)
                    time.sleep(0.1)  # 等待文件系統釋放
                except Exception as e:
                    print(f"[WARN] 無法處理舊 baseline 文件: {e}")
                    if backup_file and os.path.exists(backup_file):
                        os.remove(backup_file)
                    raise
            
            # 移動臨時文件到目標位置
            shutil.move(temp_file, baseline_file)
            
            # 清理備份文件
            if backup_file and os.path.exists(backup_file):
                os.remove(backup_file)
            
            # 強制釋放所有文件句柄
            gc.collect()
            time.sleep(0.1)
            
            print(f"[DEBUG] Baseline 保存成功: {os.path.basename(baseline_file)} (嘗試 {attempt + 1}/{max_retries})")
            return True
            
        except Exception as e:
            print(f"[WARN] Baseline 保存失敗 (嘗試 {attempt + 1}/{max_retries}): {e}")
            
            # 清理所有臨時文件
            if temp_file and os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except Exception:
                    pass
            
            # 清理可能的備份文件
            if 'backup_file' in locals() and backup_file and os.path.exists(backup_file):
                try:
                    if os.path.exists(baseline_file):
                        os.remove(baseline_file)
                    shutil.move(backup_file, baseline_file)
                except Exception:
                    pass
            
            # 強制垃圾回收和等待
            gc.collect()
            
            if attempt < max_retries - 1:
                delay = base_delay * (2 ** attempt)  # 指數退避
                print(f"[INFO] 等待 {delay} 秒後重試...")
                time.sleep(delay)
            else:
                print(f"[ERROR] 所有嘗試都失敗，無法保存 baseline: {baseline_file}")
                return False
    
    return False

def create_baseline_for_files_robust(xlsx_files, skip_force_baseline=True):
    """
    為多個檔案建立基準線
    """
    total = len(xlsx_files)
    if total == 0:
        print("[INFO] 沒有需要 baseline 的檔案。")
        settings.baseline_completed = True
        return
    
    print("\n" + "="*90 + "\n" + " BASELINE 建立程序 ".center(90, "=") + "\n" + "="*90)
    
    progress = load_progress()
    start_index = 0
    
    if progress and settings.ENABLE_RESUME:
        print(f"🔄 發現之前的進度記錄: 完成 {progress.get('completed', 0)}/{progress.get('total', 0)}")
        if input("是否要從上次中斷的地方繼續? (y/n): ").strip().lower() == 'y':
            start_index = progress.get('completed', 0)
    
    # 啟動超時處理
    if settings.ENABLE_TIMEOUT:
        from utils.helpers import timeout_handler
        timeout_thread = threading.Thread(target=timeout_handler, daemon=True)
        timeout_thread.start()
        print(f"⏰ 啟用超時保護: {settings.FILE_TIMEOUT_SECONDS} 秒")
    
    if settings.ENABLE_MEMORY_MONITOR: 
        print(f"💾 啟用記憶體監控: {settings.MEMORY_LIMIT_MB} MB")
    
    print(f"🚀 啟用優化: {[opt for flag, opt in [(settings.USE_LOCAL_CACHE, '本地緩存'), (settings.ENABLE_FAST_MODE, '快速模式')] if flag]}")
    print(f"📂 Baseline 儲存位置: {os.path.abspath(settings.LOG_FOLDER)}")
    
    if settings.USE_LOCAL_CACHE: 
        print(f"💾 本地緩存位置: {os.path.abspath(settings.CACHE_FOLDER)}")
    
    print(f"📋 要處理的檔案: {total} 個 (從第 {start_index + 1} 個開始)")
    print(f"⏰ 開始時間: {datetime.now():%Y-%m-%d %H:%M:%S}\n" + "-"*90)
    
    os.makedirs(settings.LOG_FOLDER, exist_ok=True)
    if settings.USE_LOCAL_CACHE: 
        os.makedirs(settings.CACHE_FOLDER, exist_ok=True)
    
    success_count, skip_count, error_count = 0, 0, 0
    start_time = time.time()
    
    for i in range(start_index, total):
        if settings.force_stop:
            print("\n🛑 收到停止信號，正在安全退出...")
            save_progress(i, total)
            break
        
        file_path = xlsx_files[i]
        base_name = os.path.basename(file_path)
        
        if check_memory_limit():
            print(f"⚠️ 記憶體使用量過高，暫停10秒...")
            time.sleep(10)
            if check_memory_limit(): 
                print(f"❌ 記憶體仍然過高，停止處理")
                save_progress(i, total)
                break

        file_start_time = time.time()
        print(f"[{i+1:>2}/{total}] 處理中: {base_name} (記憶體: {get_memory_usage():.1f}MB)")
        
        cell_data = None
        try:
            baseline_file = baseline_file_path(base_name)
            old_baseline = load_baseline(baseline_file)
            old_hash = old_baseline['content_hash'] if old_baseline and 'content_hash' in old_baseline else None
            
            cell_data = dump_excel_cells_with_timeout(file_path)
            
            if cell_data is None:
                if settings.current_processing_file is None and (time.time() - file_start_time) > settings.FILE_TIMEOUT_SECONDS:
                     print(f"  結果: [TIMEOUT]")
                else:
                     print(f"  結果: [READ_ERROR]")
                error_count += 1
            else:
                curr_hash = hash_excel_content(cell_data)
                if old_hash == curr_hash and old_hash is not None:
                    print(f"  結果: [SKIP] (Hash unchanged)")
                    skip_count += 1
                else:
                    curr_author = get_excel_last_author(file_path)
                    if save_baseline(baseline_file, {"last_author": curr_author, "content_hash": curr_hash, "cells": cell_data}):
                        print(f"  結果: [OK]")
                        print(f"  Baseline: {os.path.basename(baseline_file)}")
                        success_count += 1
                    else:
                        print(f"  結果: [SAVE_ERROR]")
                        error_count += 1
            
            print(f"  耗時: {time.time() - file_start_time:.2f} 秒\n")
            save_progress(i + 1, total)
            
        except Exception as e:
            print(f"  結果: [UNEXPECTED_ERROR]\n  錯誤: {e}\n  耗時: {time.time() - file_start_time:.2f} 秒\n")
            error_count += 1
            save_progress(i + 1, total)
        finally:
            if cell_data is not None: 
                del cell_data
            if 'old_baseline' in locals() and old_baseline is not None: 
                del old_baseline
            gc.collect()

    settings.baseline_completed = True
    print("-" * 90 + f"\n🎯 BASELINE 建立完成! (總耗時: {time.time() - start_time:.2f} 秒)")
    print(f"✅ 成功: {success_count}, ⏭️  跳過: {skip_count}, ❌ 失敗: {error_count}")
    
    if settings.ENABLE_RESUME and os.path.exists(settings.RESUME_LOG_FILE):
        try: 
            os.remove(settings.RESUME_LOG_FILE)
            print(f"🧹 清理進度檔案")
        except Exception: 
            pass
    
    print("\n" + "=" * 90 + "\n")
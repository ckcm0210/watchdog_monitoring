import os
import time
import threading
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import config.settings as settings
import logging

class ActivePollingHandler:
    """
    主動輪詢處理器，處理文件變更後的持續監控
    """
    def __init__(self):
        self.polling_tasks = {}
        self.lock = threading.Lock()
        self.stop_event = threading.Event()

    def start_polling(self, file_path, event_number):
        """
        根據檔案大小決定輪詢策略
        """
        try:
            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
        except (FileNotFoundError, PermissionError, OSError) as e:
            logging.warning(f"獲取檔案大小失敗: {file_path}, 錯誤: {e}")
            file_size_mb = 0
            
        if file_size_mb < settings.POLLING_SIZE_THRESHOLD_MB:
            print(f"[輪詢] 檔案: {os.path.basename(file_path)}（細file，密集輪詢，每{settings.DENSE_POLLING_INTERVAL_SEC}s，共{settings.DENSE_POLLING_DURATION_SEC}s）")
            self._start_dense_polling(file_path, event_number)
        else:
            print(f"[輪詢] 檔案: {os.path.basename(file_path)}（大file，冷靜期輪詢，每{settings.SPARSE_POLLING_INTERVAL_SEC}s）")
            self._start_sparse_polling(file_path, event_number)

    def _start_dense_polling(self, file_path, event_number):
        """
        開始密集輪詢（小檔案）
        """
        with self.lock:
            if file_path in self.polling_tasks:
                self.polling_tasks[file_path]['timer'].cancel()
                
            def task_wrapper(remaining_duration):
                self._poll_dense(file_path, event_number, remaining_duration)
                
            timer = threading.Timer(settings.DENSE_POLLING_INTERVAL_SEC, task_wrapper, args=(settings.DENSE_POLLING_DURATION_SEC,))
            self.polling_tasks[file_path] = {'timer': timer, 'remaining_duration': settings.DENSE_POLLING_DURATION_SEC}
            timer.start()
            print(f"    [輪詢啟動] {os.path.basename(file_path)}")

    def _poll_dense(self, file_path, event_number, remaining_duration):
        """
        執行密集輪詢
        """
        if self.stop_event.is_set(): 
            return
            
        print(f"    [輪詢倒數] {os.path.basename(file_path)}，尚餘: {remaining_duration}s")
        
        # 🔥 設定事件編號並執行比較
        from core.comparison import compare_excel_changes, set_current_event_number
        set_current_event_number(event_number)
        has_changes = compare_excel_changes(file_path, silent=False, event_number=event_number, is_polling=True)
        
        with self.lock:
            if file_path not in self.polling_tasks: 
                return
                
            if has_changes:
                self.polling_tasks[file_path]['remaining_duration'] = settings.DENSE_POLLING_DURATION_SEC
            else:
                self.polling_tasks[file_path]['remaining_duration'] -= settings.DENSE_POLLING_INTERVAL_SEC
                
            new_remaining_duration = self.polling_tasks[file_path]['remaining_duration']
            
            if new_remaining_duration > 0:
                def task_wrapper(): 
                    self._poll_dense(file_path, event_number, new_remaining_duration)
                new_timer = threading.Timer(settings.DENSE_POLLING_INTERVAL_SEC, task_wrapper)
                self.polling_tasks[file_path]['timer'] = new_timer
                new_timer.start()
            else:
                print(f"    [輪詢結束] {os.path.basename(file_path)}")
                self.polling_tasks.pop(file_path, None)

    def _start_sparse_polling(self, file_path, event_number):
        """
        開始稀疏輪詢（大檔案）
        """
        with self.lock:
            if file_path in self.polling_tasks:
                self.polling_tasks[file_path]['timer'].cancel()
                
            def task_wrapper():
                self._poll_sparse(file_path, event_number)
                
            timer = threading.Timer(settings.SPARSE_POLLING_INTERVAL_SEC, task_wrapper)
            self.polling_tasks[file_path] = {'timer': timer, 'waiting': True}
            timer.start()
            print(f"    [冷靜期啟動] {os.path.basename(file_path)}")

    def _poll_sparse(self, file_path, event_number):
        """
        執行稀疏輪詢
        """
        if self.stop_event.is_set(): 
            return
            
        print(f"    [冷靜期檢查] {os.path.basename(file_path)}")
        
        # 🔥 設定事件編號並執行比較
        from core.comparison import compare_excel_changes, set_current_event_number
        set_current_event_number(event_number)
        has_changes = compare_excel_changes(file_path, silent=False, event_number=event_number, is_polling=True)
        
        with self.lock:
            if file_path not in self.polling_tasks: 
                return
                
            if has_changes:
                def task_wrapper():
                    self._poll_sparse(file_path, event_number)
                new_timer = threading.Timer(settings.SPARSE_POLLING_INTERVAL_SEC, task_wrapper)
                self.polling_tasks[file_path]['timer'] = new_timer
                new_timer.start()
            else:
                print(f"    [冷靜期結束] {os.path.basename(file_path)}")
                self.polling_tasks.pop(file_path, None)

    def stop(self):
        """
        停止所有輪詢任務
        """
        self.stop_event.set()
        with self.lock:
            for task in self.polling_tasks.values(): 
                task['timer'].cancel()
            self.polling_tasks.clear()

class ExcelFileEventHandler(FileSystemEventHandler):
    """
    Excel 檔案事件處理器
    """
    def __init__(self, polling_handler):
        self.polling_handler = polling_handler
        self.last_event_times = {}
        self.event_counter = 0
        
    def on_created(self, event):
        """
        檔案建立事件處理
        """
        if event.is_directory:
            return

        file_path = event.src_path

        # [最終修正] 在處理前，先確認檔案是否仍然存在。
        # 這可以處理「建立後立刻重新命名」的競爭條件，避免處理一個已不存在的檔案。
        time.sleep(0.1) # 短暫等待，以確保 move 事件能被作業系統處理
        if not os.path.exists(file_path):
            # print(f"[DEBUG] 檔案 {os.path.basename(file_path)} 在處理前已消失，可能已被重新命名，跳過 on_created。")
            return

        # 檢查是否為支援的 Excel 檔案
        if not file_path.lower().endswith(settings.SUPPORTED_EXTS):
            return

        # 檢查是否為臨時檔案
        if os.path.basename(file_path).startswith('~$'):
            return
            
        print(f"\n✨ 發現新檔案: {os.path.basename(file_path)}")
        print(f"📊 正在建立基準線...")

        from core.baseline import create_baseline_for_files_robust
        create_baseline_for_files_robust([file_path])

        print(f"✅ 基準線建立完成，已納入監控: {os.path.basename(file_path)}")

    def on_modified(self, event):
        """
        檔案修改事件處理
        """
        if event.is_directory:
            return
            
        file_path = event.src_path
        
        # 檢查是否為支援的 Excel 檔案
        if not file_path.lower().endswith(settings.SUPPORTED_EXTS):
            return
            
        # 檢查是否為臨時檔案
        if os.path.basename(file_path).startswith('~$'):
            return
            
        # 防抖動處理
        current_time = time.time()
        if file_path in self.last_event_times:
            if current_time - self.last_event_times[file_path] < settings.DEBOUNCE_INTERVAL_SEC:
                return
                
        self.last_event_times[file_path] = current_time
        self.event_counter += 1
        
        # 獲取檔案最後作者
        try:
            from core.excel_parser import get_excel_last_author
            last_author = get_excel_last_author(file_path)
            author_info = f" (最後儲存者: {last_author})" if last_author != 'Unknown' else ""
        except Exception as e:
            author_info = ""
        
        print(f"\n🔔 檔案變更偵測: {os.path.basename(file_path)} (事件 #{self.event_counter}){author_info}")
        
        # 🔥 設定事件編號並立即執行一次比較
        from core.comparison import compare_excel_changes, set_current_event_number
        set_current_event_number(self.event_counter)
        
        print(f"📊 立即檢查變更...")
        has_changes = compare_excel_changes(file_path, silent=False, event_number=self.event_counter, is_polling=False)
        
        if has_changes:
            print(f"✅ 發現變更，開始輪詢監控...")
        else:
            print(f"ℹ️  暫未發現變更，開始輪詢監控...")
        
        # 開始輪詢
        self.polling_handler.start_polling(file_path, self.event_counter)

    def on_moved(self, event):
        """
        檔案/資料夾被移動或重新命名時觸發
        """
        if event.is_directory:
            return

        # 我們只關心重新命名後的檔案是否為 Excel 檔案
        if not event.dest_path.lower().endswith(settings.SUPPORTED_EXTS):
            return

        print(f"\n➡️  偵測到檔案重新命名/移動:")
        print(f"    來源: {os.path.basename(event.src_path)}")
        print(f"    目的: {os.path.basename(event.dest_path)}")

        from core.baseline import baseline_file_path, create_baseline_for_files_robust
        
        # 檢查舊檔案是否有對應的基準線
        src_base_name = os.path.basename(event.src_path)
        src_b_path = baseline_file_path(src_base_name)

        if os.path.exists(src_b_path):
            # 如果舊基準線存在，表示這是一個對已監控檔案的重新命名
            dest_base_name = os.path.basename(event.dest_path)
            dest_b_path = baseline_file_path(dest_base_name)
            try:
                os.rename(src_b_path, dest_b_path)
                print(f"✅ 基準線已同步更新: {os.path.basename(src_b_path)} -> {os.path.basename(dest_b_path)}")
            except OSError as e:
                print(f"❌ 更新基準線名稱失敗: {e}")
        else:
            # 如果舊基準線不存在，這很可能就是「右鍵->新增->重新命名」的流程
            # 我們將新的檔案視為一個全新的檔案來建立基準線
            print(f"ℹ️  偵測到新檔案，正在建立基準線: {os.path.basename(event.dest_path)}")
            create_baseline_for_files_robust([event.dest_path])
            print(f"✅ 基準線建立完成: {os.path.basename(event.dest_path)}")

# 創建全局輪詢處理器實例
active_polling_handler = ActivePollingHandler()



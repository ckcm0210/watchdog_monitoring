"""
系統配置設定
所有原始配置都在這裡，確保向後相容
"""
import os
from datetime import datetime

# =========== User Config ============
ENABLE_BLACK_CONSOLE = True
CONSOLE_POPUP_ON_COMPARISON = True

SCAN_ALL_MODE = True
USE_LOCAL_CACHE = True
CACHE_FOLDER = r"C:\Users\user\Desktop\watchdog\cache_folder"
ENABLE_FAST_MODE = True
ENABLE_TIMEOUT = True
FILE_TIMEOUT_SECONDS = 120
ENABLE_MEMORY_MONITOR = True
MEMORY_LIMIT_MB = 2048
ENABLE_RESUME = True
FORMULA_ONLY_MODE = True
DEBOUNCE_INTERVAL_SEC = 2

RESUME_LOG_FILE = r"C:\Users\user\Desktop\watchdog\resume_log\baseline_progress.log"
WATCH_FOLDERS = [
    r"C:\Users\user\Desktop\Test",
]
MANUAL_BASELINE_TARGET = []
LOG_FOLDER = r"C:\Users\user\Desktop\watchdog\log_folder"
LOG_FILE_DATE = datetime.now().strftime('%Y%m%d')
CSV_LOG_FILE = os.path.join(LOG_FOLDER, f"excel_change_log_{LOG_FILE_DATE}.csv.gz")
SUPPORTED_EXTS = ('.xlsx', '.xlsm')
MAX_RETRY = 10
RETRY_INTERVAL_SEC = 2
USE_TEMP_COPY = True
WHITELIST_USERS = ['ckcm0210', 'yourwhiteuser']
LOG_WHITELIST_USER_CHANGE = True
FORCE_BASELINE_ON_FIRST_SEEN = [
    r"\\network_drive\\your_folder1\\must_first_baseline.xlsx",
    "force_this_file.xlsx"
]

# =========== Polling Config ============
POLLING_SIZE_THRESHOLD_MB = 10
DENSE_POLLING_INTERVAL_SEC = 5
DENSE_POLLING_DURATION_SEC = 15
SPARSE_POLLING_INTERVAL_SEC = 15
SPARSE_POLLING_DURATION_SEC = 15

# =========== 全局變數 ============
current_processing_file = None
processing_start_time = None
force_stop = False
baseline_completed = False

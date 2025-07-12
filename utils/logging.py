"""
日誌和打印功能
"""
import builtins
import logging
import os
from datetime import datetime
from io import StringIO
from wcwidth import wcswidth, wcwidth

# 保存原始 print 函數
_original_print = builtins.print

def timestamped_print(*args, **kwargs):
    """
    帶時間戳的打印函數
    """
    # 如果有 file=... 參數，直接用原生 print
    if 'file' in kwargs:
        _original_print(*args, **kwargs)
        return

    output_buffer = StringIO()
    _original_print(*args, file=output_buffer, **kwargs)
    message = output_buffer.getvalue()
    output_buffer.close()

    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    # 簡化邏輯：所有行都加時間戳記
    lines = message.rstrip().split('\n')
    timestamped_lines = []
    
    for line in lines:
        timestamped_lines.append(f"[{timestamp}] {line}")
    
    timestamped_message = '\n'.join(timestamped_lines)
    _original_print(timestamped_message)
    
    # 檢查是否為比較表格訊息
    is_comparison = any(keyword in message for keyword in [
        'Address', 'Baseline', 'Current', 
        '[SUMMARY]', '====', '----',
        '[MOD]', '[ADD]', '[DEL]'
    ])
    
    # 同時送到黑色 console - 使用延遲導入避免循環導入
    try:
        from ui.console import black_console
        if black_console and black_console.running:
            black_console.add_message(timestamped_message, is_comparison=is_comparison)
    except ImportError:
        pass

def init_logging():
    """
    初始化日誌系統
    """
    # 設置原有的時間戳打印系統
    builtins.print = timestamped_print
    
    # 設置專業級日誌系統
    setup_professional_logging()

def setup_professional_logging():
    """
    設置專業級日誌系統
    """
    # 創建自定義格式化器
    class ChineseFormatter(logging.Formatter):
        """支持中文的自定義格式化器"""
        
        def format(self, record):
            # 添加中文級別名稱
            level_names = {
                'DEBUG': '調試',
                'INFO': '信息', 
                'WARNING': '警告',
                'ERROR': '錯誤',
                'CRITICAL': '嚴重'
            }
            
            record.level_zh = level_names.get(record.levelname, record.levelname)
            return super().format(record)
    
    # 獲取根日誌器
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    
    # 避免重複添加處理器
    if logger.handlers:
        return
    
    # 創建控制台處理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    
    # 設置格式
    formatter = ChineseFormatter(
        '[%(asctime)s] [%(level_zh)s] %(name)s: %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    console_handler.setFormatter(formatter)
    
    # 添加處理器
    logger.addHandler(console_handler)

def get_logger(name=None):
    """
    獲取日誌器實例
    
    Args:
        name: 日誌器名稱，通常使用 __name__
        
    Returns:
        logging.Logger: 日誌器實例
    """
    return logging.getLogger(name or 'watchdog_monitoring')

def wrap_text_with_cjk_support(text, width):
    """
    自研的、支持 CJK 字符寬度的智能文本換行函數
    """
    lines = []
    line = ""
    current_width = 0
    for char in text:
        char_width = wcwidth(char)
        if char_width < 0: 
            continue # 跳過控制字符

        if current_width + char_width > width:
            lines.append(line)
            line = char
            current_width = char_width
        else:
            line += char
            current_width += char_width
    if line:
        lines.append(line)
    return lines or ['']

def _get_display_width(text):
    """
    精準計算一個字串的顯示闊度，處理 CJK 全形字元
    """
    return wcswidth(str(text))
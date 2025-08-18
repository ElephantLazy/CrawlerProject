# 性能配置檔案
class PerformanceConfig:
    # 瀏覽器配置
    BROWSER_OPTIONS = [
        '--headless',
        '--no-sandbox',
        '--disable-dev-shm-usage',
        '--disable-gpu',
        '--disable-extensions',
        '--disable-logging',
        '--disable-web-security',
        '--disable-features=VizDisplayCompositor',
        '--disable-images',  # 不載入圖片（除了驗證碼）
        '--disable-javascript',  # 如果網站不需要JS可以禁用
        '--disable-plugins',
        '--disable-background-timer-throttling',
        '--disable-renderer-backgrounding',
        '--disable-backgrounding-occluded-windows',
        '--disable-features=TranslateUI',
        '--disable-background-networking'
    ]
    
    # 線程配置
    MAX_THREADS = 6  # 建議值，可根據系統性能調整
    BATCH_SIZE = 20  # 每次處理的記錄數量
    
    # 資料庫配置
    DB_PRAGMA_SETTINGS = [
        "PRAGMA journal_mode=WAL",
        "PRAGMA synchronous=NORMAL", 
        "PRAGMA cache_size=10000",
        "PRAGMA temp_store=memory",
        "PRAGMA mmap_size=268435456"  # 256MB
    ]
    
    # 驗證碼處理配置
    CAPTCHA_RETRY_MAX = 2
    CAPTCHA_TIMEOUT = 3
    
    # 等待時間配置
    ELEMENT_WAIT_TIME = 5
    CAPTCHA_REFRESH_WAIT = 0.3
    THREAD_START_DELAY = 0.1
    
    # OCR配置
    TESSERACT_CONFIG = '--psm 8 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'
    
    # 圖像處理配置
    IMAGE_SCALE_FACTOR = 3
    BINARY_THRESHOLD = 140

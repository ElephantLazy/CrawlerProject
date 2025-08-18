# 性能監控工具（簡化版本）
import time
import threading
from datetime import datetime

class PerformanceMonitor:
    def __init__(self):
        self.start_time = time.time()
        self.processed_count = 0
        self.success_count = 0
        self.error_count = 0
        self.lock = threading.Lock()
        
    def record_processed(self):
        with self.lock:
            self.processed_count += 1
            
    def record_success(self):
        with self.lock:
            self.success_count += 1
            
    def record_error(self):
        with self.lock:
            self.error_count += 1
    
    def get_stats(self):
        with self.lock:
            elapsed_time = time.time() - self.start_time
            rate = self.processed_count / elapsed_time if elapsed_time > 0 else 0
            success_rate = self.success_count / self.processed_count if self.processed_count > 0 else 0
            
            return {
                'elapsed_time': elapsed_time,
                'processed': self.processed_count,
                'success': self.success_count,
                'errors': self.error_count,
                'rate_per_sec': rate,
                'success_rate': success_rate
            }
    
    def print_stats(self):
        stats = self.get_stats()
        print(f"\n=== 性能統計 ===")
        print(f"運行時間: {stats['elapsed_time']:.1f} 秒")
        print(f"已處理: {stats['processed']} 筆")
        print(f"成功: {stats['success']} 筆")
        print(f"錯誤: {stats['errors']} 筆")
        print(f"處理速度: {stats['rate_per_sec']:.2f} 筆/秒")
        print(f"成功率: {stats['success_rate']:.2%}")
        print("================\n")

# 全域監控實例
monitor = PerformanceMonitor()

def start_monitoring():
    """啟動性能監控線程"""
    def monitor_loop():
        while True:
            time.sleep(60)  # 每分鐘顯示一次統計
            monitor.print_stats()
    
    monitor_thread = threading.Thread(target=monitor_loop, daemon=True)
    monitor_thread.start()

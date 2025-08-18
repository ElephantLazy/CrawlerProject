#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tesseract 訓練器使用示例

這個腳本展示如何使用訓練器的各個功能
"""

from tesseract_trainer import TesseractTrainer
import time

def demo_workflow():
    """示範完整的訓練工作流程"""
    print("🎯 Tesseract 訓練器工作流程示範")
    print("="*50)
    
    # 初始化訓練器
    trainer = TesseractTrainer()
    
    print("\n📊 當前狀態:")
    trainer.show_statistics()
    
    print("\n💡 接下來您可以：")
    print("1. 運行 python tesseract_trainer.py 開始互動式訓練")
    print("2. 或者使用 start_tesseract_trainer.bat 快速啟動")
    
    print("\n📋 建議的工作流程:")
    print("1️⃣  收集 100-200 張驗證碼圖片")
    print("2️⃣  手動標註每張圖片的正確答案")
    print("3️⃣  生成 Tesseract 訓練數據")
    print("4️⃣  測試模型效果")
    print("5️⃣  根據結果決定是否需要更多數據")

def quick_test():
    """快速測試功能"""
    print("🧪 快速測試模式")
    print("-"*30)
    
    trainer = TesseractTrainer()
    
    # 檢查是否有現有數據
    if trainer.annotations:
        print(f"✅ 發現 {len(trainer.annotations)} 個已標註的圖片")
        
        # 運行測試
        trainer.test_current_model()
    else:
        print("📝 尚未有標註數據，請先進行以下步驟：")
        print("1. 收集驗證碼圖片")
        print("2. 手動標註正確答案")

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == "test":
        quick_test()
    else:
        demo_workflow()

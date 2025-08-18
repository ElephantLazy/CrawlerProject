# 🎯 Tesseract-OCR 驗證碼訓練器 - 完整解決方案

## 📦 包含的文件

```
📁 CrawlerProject/
├── 🔧 tesseract_trainer.py          # 主要訓練器程式
├── 📖 TESSERACT_TRAINING_GUIDE.md   # 詳細使用指南
├── 🚀 start_tesseract_trainer.bat   # 快速啟動腳本
├── 🧪 tesseract_demo.py             # 使用示例
├── 📁 tesseract_training/            # 訓練數據目錄（自動創建）
│   ├── images/                      # 驗證碼圖片
│   ├── annotations/                 # 標註文件
│   ├── training_data/               # Tesseract 訓練數據
│   ├── models/                      # 生成的模型
│   └── annotations.json             # 標註數據庫
└── 📄 src/20240920.py               # 原始爬蟲程式
```

## 🚀 快速開始

### 方法一：雙擊批次文件
```
雙擊 → start_tesseract_trainer.bat
```

### 方法二：命令行啟動
```bash
cd d:\workspace\CrawlerProject
python tesseract_trainer.py
```

### 方法三：查看示例
```bash
python tesseract_demo.py
```

## 📋 完整工作流程

### 1️⃣ 收集驗證碼圖片 (5-10分鐘)
- 啟動訓練器
- 選擇選項 `1`
- 輸入數量 (建議100-200張)
- 等待自動收集完成

### 2️⃣ 手動標註 (1-3小時)
- 選擇選項 `2`
- 用圖片檢視器開啟每張圖片
- 輸入正確的6位驗證碼
- 系統自動保存進度

### 3️⃣ 生成訓練數據 (1分鐘)
- 選擇選項 `3`
- 自動生成 Tesseract 訓練格式
- 創建訓練腳本

### 4️⃣ 測試效果 (1分鐘)
- 選擇選項 `5`
- 查看當前識別準確率
- 決定是否需要更多數據

## 💡 標註技巧

### 🖼️ 查看圖片：
1. **雙擊圖片** → Windows 照片檢視器
2. **拖拽到瀏覽器** → 在瀏覽器中查看
3. **右鍵 → 開啟方式 → 小畫家** → 可以放大查看

### ✏️ 輸入原則：
- ✅ 輸入6位英數字 (如: ABC123)
- ✅ 統一使用大寫
- ✅ 看不清楚時輸入 `skip`
- ❌ 不要包含空格或符號

### ⌨️ 快捷指令：
- `skip` - 跳過當前圖片
- `save` - 手動保存進度
- `quit` - 退出標註模式

## 📊 預期效果

| 標註數量 | 預期準確率 | 建議用途 |
|---------|-----------|---------|
| 50-100張 | 60-70% | 初步測試 |
| 100-200張 | 70-80% | 基本可用 |
| 200-500張 | 80-90% | 生產環境 |
| 500+張 | 90%+ | 高精度需求 |

## 🔧 整合到主程式

訓練完成後，在主程式中使用：

```python
# 方法一：使用改進的配置
result = pytesseract.image_to_string(
    img, 
    config='--psm 8 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'
)

# 方法二：如果生成了自定義模型
result = pytesseract.image_to_string(
    img, 
    lang='captcha',  # 自定義模型名稱
    config='--psm 8'
)
```

## 🎯 最佳實踐

### 🖼️ 圖片品質：
- 確保收集的圖片清晰
- 包含各種不同的驗證碼類型
- 避免重複的驗證碼

### 📝 標註品質：
- 仔細檢查每個字符
- 統一大小寫處理
- 跳過無法辨識的圖片

### 🧪 測試策略：
- 定期測試模型效果
- 收集錯誤樣本進行補充標註
- 調整 Tesseract 參數配置

## ❓ 疑難排解

### Q: 程式無法啟動？
A: 檢查 Python 和必要套件是否安裝：
```bash
pip install selenium pillow pytesseract requests
```

### Q: 無法下載驗證碼？
A: 
1. 檢查網路連線
2. 確認 Chrome 和 ChromeDriver 版本匹配
3. 檢查防火牆設置

### Q: 標註進度丟失？
A: 
1. 訓練器每10張自動保存
2. 可手動輸入 `save` 保存
3. 數據保存在 `annotations.json`

### Q: 識別效果不佳？
A: 
1. 增加標註數據量
2. 檢查標註準確性
3. 嘗試不同的 Tesseract 參數

## 🔮 進階功能

### 自動化標註 (實驗性)：
可以結合現有的識別結果進行半自動標註，加快標註速度。

### 圖片預處理：
可以添加更多圖片處理步驟來提高訓練數據品質。

### 模型評估：
可以添加更詳細的模型評估指標，如字符級別的準確率。

---

## 🎉 開始您的 Tesseract 訓練之旅！

現在您已經有了完整的工具和指南，可以開始創建屬於您自己的高精度驗證碼識別模型了！

**祝您訓練順利！** 🚀

# Tesseract-OCR 驗證碼訓練器使用指南

## 🎯 功能概述

這個訓練器專門用於收集驗證碼圖片、手動標註正確答案，並生成 Tesseract-OCR 自定義訓練模型。

## 📋 系統需求

### 必需軟體：
1. **Python 3.7+** 
2. **Tesseract-OCR** (已安裝在 `C:\Program Files\Tesseract-OCR\`)
3. **Chrome 瀏覽器**
4. **ChromeDriver**

### Python 套件：
```bash
pip install selenium pillow pytesseract requests
```

### Tesseract 訓練工具（進階）：
如果要訓練自定義模型，需要額外下載 Tesseract 訓練工具。

## 🚀 使用步驟

### 第一步：啟動訓練器
```bash
cd d:\workspace\CrawlerProject
python tesseract_trainer.py
```

### 第二步：收集驗證碼圖片
1. 選擇選項 `1` - 收集驗證碼圖片
2. 輸入要收集的數量（建議 100-500 張）
3. 等待自動收集完成

### 第三步：手動標註
1. 選擇選項 `2` - 手動標註驗證碼
2. 系統會提示您開啟圖片文件
3. 查看圖片後輸入正確的 6 位驗證碼
4. 繼續標註所有圖片

### 第四步：生成訓練數據
1. 選擇選項 `3` - 生成訓練數據
2. 系統會自動創建 Tesseract 訓練格式文件
3. 生成訓練腳本

### 第五步：測試效果
1. 選擇選項 `5` - 測試當前模型
2. 查看識別準確率

## 📁 文件結構

```
tesseract_training/
├── images/                    # 原始驗證碼圖片
│   ├── captcha_20250806_143000_0001.png
│   ├── captcha_20250806_143001_0002.png
│   └── ...
├── annotations/              # 標註相關文件
├── training_data/           # Tesseract 訓練數據
│   ├── captcha.font.exp0000.tif
│   ├── captcha.font.exp0000.box
│   ├── train_model.bat      # 訓練腳本
│   └── ...
├── models/                  # 生成的模型文件
├── annotations.json         # 標註數據庫
└── ...
```

## 💡 標註技巧

### 查看圖片的方法：
1. **Windows 照片檢視器**：雙擊圖片文件
2. **畫圖軟體**：右鍵 → 開啟方式 → 小畫家
3. **瀏覽器**：拖拽圖片到瀏覽器窗口
4. **圖片編輯軟體**：如 GIMP、Photoshop 等

### 標註原則：
- ✅ 輸入完整的 6 位驗證碼
- ✅ 區分大小寫（建議統一用大寫）
- ✅ 只輸入英數字符 (0-9, A-Z)
- ❌ 不要包含空格或特殊符號
- ❌ 如果圖片模糊無法辨識，輸入 `skip` 跳過

### 快捷操作：
- `skip` - 跳過當前圖片
- `quit` - 退出標註模式
- `save` - 手動保存進度
- 系統每標註 10 張會自動保存

## 🧪 測試模型效果

運行測試功能可以查看當前 Tesseract 的識別效果：

```
測試結果示例：
✅ captcha_001.png: 'ABC123' == 'ABC123'
❌ captcha_002.png: 'XYZ789' != 'XY2789'
📊 測試結果: 8/10 正確，準確率: 80.0%
```

## 🏗️ 訓練自定義模型（進階）

### 安裝訓練工具：
1. 下載 Tesseract 訓練工具
2. 解壓到 Tesseract 安裝目錄

### 運行訓練：
1. 進入 `training_data` 目錄
2. 運行 `train_model.bat`
3. 等待訓練完成
4. 將生成的 `.traineddata` 文件複製到 Tesseract 的 `tessdata` 目錄

### 使用自定義模型：
```python
# 在主程式中使用自定義模型
result = pytesseract.image_to_string(
    img, 
    lang='captcha',  # 使用自定義模型
    config='--psm 8'
)
```

## 📊 建議的數據量

| 目標準確率 | 建議圖片數量 | 標註時間 |
|-----------|-------------|----------|
| 70-80%    | 100-200張   | 1-2小時  |
| 80-90%    | 300-500張   | 3-5小時  |
| 90%+      | 500-1000張  | 5-10小時 |

## ❓ 常見問題

### Q: 無法下載驗證碼圖片？
A: 檢查網路連線和 Chrome 設置，確保能正常訪問財政部網站。

### Q: 標註時看不清圖片內容？
A: 可以用圖片編輯軟體放大查看，或輸入 `skip` 跳過模糊圖片。

### Q: 訓練腳本執行失敗？
A: 需要安裝完整的 Tesseract 訓練工具，或聯繫技術支援。

### Q: 如何提高識別準確率？
A: 
1. 增加標註數據量
2. 確保標註準確性
3. 對圖片進行預處理
4. 調整 Tesseract 配置參數

## 🔧 進階配置

### 圖片預處理：
可以修改 `create_training_data()` 函數來添加：
- 圖片去噪
- 對比度增強
- 尺寸標準化

### Tesseract 參數調優：
在主程式中嘗試不同的 PSM 模式和 OEM 引擎：
```python
configs = [
    '--psm 6 --oem 3',
    '--psm 7 --oem 1', 
    '--psm 8 --oem 3',
    '--psm 13'
]
```

## 📞 技術支援

如果遇到問題，請檢查：
1. Python 環境和套件安裝
2. Tesseract 安裝路徑
3. Chrome 和 ChromeDriver 版本兼容性
4. 網路連線狀況

---

**祝您訓練順利！** 🎉

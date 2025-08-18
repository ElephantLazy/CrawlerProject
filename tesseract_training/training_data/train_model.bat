@echo off
REM Tesseract 訓練腳本
REM 請確保已安裝 Tesseract 和訓練工具

echo 開始 Tesseract 訓練...

REM 設置語言名稱
set LANG_NAME=captcha

REM 合併所有 BOX 和 TIF 文件
echo 合併訓練文件...
copy /b *.box %LANG_NAME%.box
for %%f in (*.tif) do (
    echo 處理 %%f...
    tesseract %%f %%~nf batch.nochop makebox
)

REM 生成訓練文件
echo 生成訓練文件...
tesseract %LANG_NAME%.font.exp0000.tif %LANG_NAME%.font.exp0000 batch.nochop makebox
for %%f in (*.tif) do tesseract %%f %%~nf box.train

REM 計算字符集
echo 計算字符集...
unicharset_extractor *.box

REM 生成字形統計
echo 生成字形統計...
shapeclustering -F font_properties -U unicharset *.tr

REM 生成聚類
echo 生成聚類...
mftraining -F font_properties -U unicharset -O %LANG_NAME%.unicharset *.tr

REM 生成標準化特徵文件
echo 生成標準化特徵文件...
cntraining *.tr

REM 重命名文件
echo 重命名輸出文件...
ren normproto %LANG_NAME%.normproto
ren inttemp %LANG_NAME%.inttemp
ren pffmtable %LANG_NAME%.pffmtable
ren shapetable %LANG_NAME%.shapetable

REM 合併訓練文件
echo 合併最終模型...
combine_tessdata %LANG_NAME%.

echo 訓練完成！生成的模型文件: %LANG_NAME%.traineddata
echo 請將此文件複製到 Tesseract 的 tessdata 目錄中
pause

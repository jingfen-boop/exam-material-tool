# 會考教材產製工具 Web v1.0

這是一個 Streamlit 網頁版原型，目標流程：

1. 上傳「會考國文題本 PDF」
2. 上傳「官方答案 PDF」
3. 上傳「各題通過率 PDF」
4. 自動配對題號、答案、通過率
5. 依通過率篩題與手動勾選
6. 編輯能力類型、解析、教學步驟、筆記策略
7. 產生學生題本 Word 與詳解／教學法 Word
8. 圖片、圖表或複雜版面可選擇保留原 PDF 題目裁圖

## 最快部署：Streamlit Community Cloud

把本資料夾內容上傳到一個 GitHub repository，然後在 Streamlit Community Cloud 建立 App：
- Main file path：`app.py`
- Python requirements：平台會自動讀取 `requirements.txt`

部署完成後會得到一個 HTTPS 網址，可直接用瀏覽器開啟。

## 本機測試

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

瀏覽器會開啟本機網頁介面。

## AI 功能

AI 是選用功能。不輸入 API Key 時，題庫建立、通過率篩選、Word 輸出都可以正常使用。
若輸入 OpenAI API Key，可針對單題產生：
- 能力類型
- 逐選項解析
- 教學步驟
- 筆記策略

AI 產出屬初稿，請人工審核。

## v1.0 已知限制

- 不同年度 PDF 的排版差異可能造成少數題目拆題不完整，需人工校對。
- 複雜圖像題預設以原 PDF 裁圖保留，確保版面與資訊不遺失。
- 題組目前可辨識「請閱讀…並回答 X～Y 題」範圍，並在篩選時整組保留；共用選文尚未獨立抽成可編輯區塊。
- Word 版面以目前提供的範例風格重建，尚不是像素級複製。

## 建議下一版

- 題組共用選文獨立化
- 題目圖片拆成獨立圖片物件
- 可編輯 Word 文字與原版面更精準對齊
- 題庫持久化（SQLite / 雲端資料庫）
- 登入與多人協作

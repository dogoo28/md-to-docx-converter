# 專業級 Markdown 轉 Docx 工具 (Enhanced Converter)

這是一個基於 Streamlit、Pandoc 與 Python-docx 的高級文件轉換工具，專為需要高品質 Word 輸出的使用者設計。除了基本的 Markdown 轉換外，還支援 Mermaid 流程圖、數學公式、自訂樣式範本以及多檔合併功能。

## ✨ 主要功能

### 1. 核心轉換
*   **Markdown 轉 Docx**：基於 Pandoc 強大的轉換引擎。
*   **Mermaid 圖表支援**：自動解析 Markdown 中的 ` ```mermaid` 區塊，透過 Mermaid.ink 渲染為高解析度圖片並嵌入 Word 文件中。
*   **數學公式**：支援 LaTeX 格式數學公式 ($...$)。

### 2. 多元輸入方式
*   **上傳檔案模式**：上傳一個或多個 `.md` 檔案進行轉換。
*   **貼上文字模式**：直接貼上 Markdown 內容，無需建立檔案即可快速轉換。

### 3. 樣式與範本管理
*   **內建/自訂範本**：可選擇內建的標準範本，或上傳自己的 `.docx` 作為樣式參考 (Reference Doc)。
*   **樣式預覽**：提供直觀的樣式檢視面板，可解析範本中的字型、大小、顏色，確保輸出符合預期。

### 4. 進階輸出
*   **合併輸出 (Merge)**：支援上傳多個 `.md` 檔案並合併為單一 Word 文件。
    *   **檔案排序**：提供介面調整合併順序。
    *   **智慧分頁**：每個章節之間自動插入 OpenXML 標準分頁符號。
*   **目錄生成 (TOC)**：可選擇自動產生目錄，並支援 Metadata (標題、作者、日期) 設定。
*   **單檔直接下載**：單個檔案轉換時直接下載 `.docx`，多檔才使用 ZIP 打包。

## 🚀 快速開始

### 安裝需求

請確保已安裝 [Pandoc](https://pandoc.org/installing.html) 並配置於系統路徑中。

```bash
pip install -r requirements.txt
```

### 啟動程式

```bash
streamlit run app.py
```

## 📂 專案結構

*   `app.py`: 主程式 UI 介面與互動邏輯。
*   `core/`: 核心邏輯模組。
    *   `converter.py`: Markdown 轉 Docx 與圖表處理。
    *   `style_analyzer.py`: Word 範本樣式分析。
*   `templates/`: 存放內建 Word 範本。
*   `系統文件/`: 包含操作手冊與系統相關說明。

## 📝 使用說明

詳細操作流程請參閱 [操作手冊](系統文件/操作手冊.md)。

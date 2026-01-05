import streamlit as st
import pypandoc
import os
import tempfile
import zipfile
import re
import base64
import requests
from io import BytesIO

# --- 頁面配置 ---
st.set_page_config(page_title="專業級 MD 轉 Word 工具", page_icon="📈", layout="wide")

# --- 核心功能：處理 Mermaid 並轉換為本地圖片 ---
def process_mermaid_to_local_img(md_text, tmpdir):
    """
    解析 Markdown 中的 Mermaid 區塊，將其轉換為 URL 安全的編碼，
    並下載為實體 PNG 檔案供 Pandoc 嵌入。
    """
    # 統一處理換行符號，避免匹配失敗
    md_text = md_text.replace('\r\n', '\n')
    
    # 正規表示式：匹配 ```mermaid ... ``` 區塊
    pattern = re.compile(r"```mermaid\s+(.*?)```", re.DOTALL | re.IGNORECASE)
    
    def download_img(match):
        mermaid_code = match.group(1).strip()
        if not mermaid_code:
            return ""
            
        try:
            # 修正：使用 UTF-8 編碼並改用 urlsafe_b64encode 處理中文字元與特殊符號
            code_bytes = mermaid_code.encode('utf-8')
            base64_code = base64.urlsafe_b64encode(code_bytes).decode('utf-8').replace('=', '')
            
            # 使用 mermaid.ink 的圖片渲染路徑
            url = f"https://mermaid.ink/img/{base64_code}"
            
            # 建立本地臨時圖檔路徑
            img_filename = f"chart_{os.urandom(4).hex()}.png"
            img_path = os.path.join(tmpdir, img_filename)
            
            # 執行下載，增加 timeout 以應對複雜圖表的渲染時間
            resp = requests.get(url, timeout=30)
            
            if resp.status_code == 200:
                with open(img_path, "wb") as f:
                    f.write(resp.content)
                # 重要：返回本地實體路徑，前後加上換行確保 Word 格式正確
                return f"\n\n![Mermaid Chart]({img_path})\n\n"
            else:
                st.error(f"Mermaid 渲染失敗 (HTTP {resp.status_code})。請檢查語法或網路連結。")
                return f"\n\n> [!CAUTION] Mermaid 渲染失敗 (HTTP {resp.status_code})\n\n```mermaid\n{mermaid_code}\n```\n\n"
                
        except Exception as e:
            st.warning(f"圖表下載異常: {e}")
            return f"\n\n```mermaid\n{mermaid_code}\n```\n\n"

    return pattern.sub(download_img, md_text)

# --- Streamlit UI 介面 ---
st.title("📈 專業級 Markdown 轉 Word 工具")
st.markdown("已優化：**支援中文流程圖**、**自動嵌入圖片**、**自定義 Word 範本**。")

with st.sidebar:
    st.header("⚙️ 轉換設定")
    add_toc = st.checkbox("自動生成目錄 (TOC)", value=True)
    math_support = st.checkbox("支援數學公式 ($LaTeX$)", value=True)
    
    st.divider()
    
    st.subheader("🎨 樣式範本")
    ref_file = st.file_uploader("上傳參考 Word (.docx)", type=["docx"], help="輸出的 Word 將繼承此檔案的字型與標題樣式。")
    if ref_file:
        st.success("✅ 樣式範本已載入")

# --- 檔案處理區 ---
uploaded_files = st.file_uploader("上傳一個或多個 .md 檔案", type=["md"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 開始批次轉換", use_container_width=True):
        zip_buffer = BytesIO()
        
        try:
            with st.spinner('正在分析架構圖並轉換格式...'):
                # 使用臨時資料夾存放中間產物
                with tempfile.TemporaryDirectory() as tmpdir:
                    
                    # 處理範本路徑
                    ref_path = None
                    if ref_file:
                        ref_path = os.path.join(tmpdir, "user_style.docx")
                        with open(ref_path, "wb") as f:
                            f.write(ref_file.getbuffer())

                    # 建立 ZIP 檔案
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        progress_bar = st.progress(0)
                        
                        for i, file in enumerate(uploaded_files):
                            # 讀取 MD 原始文字
                            raw_md = file.read().decode("utf-8")
                            
                            # 1. 預處理：將 Mermaid 代碼轉換為本地 PNG 圖片路徑
                            processed_md = process_mermaid_to_local_img(raw_md, tmpdir)
                            
                            output_filename = file.name.replace(".md", ".docx")
                            temp_docx_path = os.path.join(tmpdir, f"output_{i}.docx")
                            
                            # 2. 設定 Pandoc 參數
                            # --extract-media=. 確保所有本地圖片被正確打包進 docx
                            args = ["--standalone", "--extract-media=."]
                            if add_toc: args.append("--toc")
                            if math_support: args.append("--mathjax")
                            if ref_path: args.append(f"--reference-doc={ref_path}")
                            
                            # 3. 執行轉換（使用實體檔案路徑）
                            pypandoc.convert_text(
                                processed_md, 
                                'docx', 
                                format='md', 
                                extra_args=args, 
                                outputfile=temp_docx_path
                            )
                            
                            # 4. 將結果讀入並打包至 ZIP
                            with open(temp_docx_path, "rb") as f:
                                zip_file.writestr(output_filename, f.read())
                            
                            progress_bar.progress((i + 1) / len(uploaded_files))

            st.success("✨ 轉換成功！所有圖表已嵌入檔案中。")
            st.download_button(
                label="📥 下載轉換後的 ZIP 包",
                data=zip_buffer.getvalue(),
                file_name="converted_documents.zip",
                mime="application/zip",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"轉換失敗：{str(e)}")
            st.info("💡 提示：若出現 404，請確認您的 Mermaid 語法是否能在一般編輯器中正確顯示。")
else:
    st.info("請先上傳 Markdown 檔案開始作業。")

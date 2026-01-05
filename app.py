import streamlit as st
import pypandoc
import os
import tempfile
import zipfile
import re
import base64
import requests
from io import BytesIO

# --- 頁面設定 ---
st.set_page_config(page_title="專業 MD 轉 Word (含圖表)", page_icon="📈", layout="wide")

# --- 函式：處理 Mermaid 並下載為本地圖片 ---
def process_mermaid_to_local_img(md_text, tmpdir):
    """
    將 MD 中的 mermaid 區塊抓出來，下載成 PNG 存放在 tmpdir，
    並將 MD 內容替換為指向該本地路徑的圖片語法。
    """
    def download_img(match):
        mermaid_code = match.group(1).strip()
        try:
            # 將 Mermaid 語法編碼
            code_bytes = mermaid_code.encode('utf-8')
            base64_code = base64.b64encode(code_bytes).decode('utf-8')
            url = f"https://mermaid.ink/png/{base64_code}"
            
            # 建立本地臨時圖檔
            img_filename = f"chart_{base64.b16encode(os.urandom(4)).decode()}.png"
            img_path = os.path.join(tmpdir, img_filename)
            
            # 下載圖片
            resp = requests.get(url, timeout=15)
            if resp.status_code == 200:
                with open(img_path, "wb") as f:
                    f.write(resp.content)
                # 回傳本地路徑語法，Pandoc 才能讀取並嵌入
                return f"\n![Flowchart]({img_path})\n"
        except Exception as e:
            st.warning(f"Mermaid 轉換失敗：{e}")
        
        # 若失敗則保留原樣
        return f"\n```mermaid\n{mermaid_code}\n```\n"

    # 匹配 ```mermaid ... ```
    return re.sub(r"```mermaid\s+(.*?)\s+```", download_img, md_text, flags=re.DOTALL)

# --- UI 介面 ---
st.title("📈 專業級 Markdown 轉 Word 工具")
st.markdown("本工具支援 **Mermaid 流程圖自動嵌入**、**數學公式**及**自定義 Word 範本**。")

with st.sidebar:
    st.header("⚙️ 轉換設定")
    add_toc = st.checkbox("自動生成目錄 (TOC)", value=True)
    math_support = st.checkbox("支援數學公式 ($LaTeX$)", value=True)
    
    st.divider()
    
    st.subheader("🎨 樣式範本")
    ref_file = st.file_uploader("上傳參考 Word (.docx)", type=["docx"])
    if ref_file:
        st.success("✅ 樣式已載入")

# --- 檔案上傳 ---
uploaded_files = st.file_uploader("請上傳 Markdown 檔案 (.md)", type=["md"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 開始轉換並下載", use_container_width=True):
        zip_buffer = BytesIO()
        
        try:
            with st.spinner('正在處理圖表與轉換格式...'):
                with tempfile.TemporaryDirectory() as tmpdir:
                    
                    # 處理範本檔案
                    ref_path = None
                    if ref_file:
                        ref_path = os.path.join(tmpdir, "template.docx")
                        with open(ref_path, "wb") as f:
                            f.write(ref_file.getbuffer())

                    # 建立 ZIP
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        progress_bar = st.progress(0)
                        
                        for i, file in enumerate(uploaded_files):
                            # 讀取 MD
                            raw_md = file.read().decode("utf-8")
                            
                            # 1. 重要：預處理 Mermaid 並下載圖片到臨時資料夾
                            processed_md = process_mermaid_to_local_img(raw_md, tmpdir)
                            
                            output_filename = file.name.replace(".md", ".docx")
                            temp_docx_path = os.path.join(tmpdir, f"out_{i}.docx")
                            
                            # 2. 設定 Pandoc 參數
                            # 關鍵：--extract-media=. 確保 Pandoc 處理本地資源
                            args = ["--standalone", "--extract-media=."]
                            if add_toc: args.append("--toc")
                            if math_support: args.append("--mathjax")
                            if ref_path: args.append(f"--reference-doc={ref_path}")
                            
                            # 3. 執行轉換
                            pypandoc.convert_text(
                                processed_md, 
                                'docx', 
                                format='md', 
                                extra_args=args, 
                                outputfile=temp_docx_path
                            )
                            
                            # 4. 寫入 ZIP
                            with open(temp_docx_path, "rb") as f:
                                zip_file.writestr(output_filename, f.read())
                            
                            progress_bar.progress((i + 1) / len(uploaded_files))

            st.success("✨ 轉換成功！圖表已嵌入 Word 檔案中。")
            st.download_button(
                label="📥 下載轉換後的 ZIP 包",
                data=zip_buffer.getvalue(),
                file_name="converted_docs.zip",
                mime="application/zip",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"轉換過程中發生錯誤：{str(e)}")

else:
    st.info("請先上傳 .md 檔案。")

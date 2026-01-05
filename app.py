import streamlit as st
import pypandoc
import os
import tempfile
import zipfile
import re
import base64
import requests
from io import BytesIO

st.set_page_config(page_title="專業 MD 轉 Word (含圖表)", page_icon="📈", layout="wide")

# --- 強化版 Mermaid 處理函式 ---
def process_mermaid_to_local_img(md_text, tmpdir):
    # 匹配 ```mermaid (不論大小寫) 開頭，直到下一組 ``` 結束
    # 包含處理可能出現的 \r\n 換行符號
    pattern = re.compile(r"```mermaid\s+(.*?)```", re.DOTALL | re.IGNORECASE)
    
    def download_img(match):
        mermaid_code = match.group(1).strip()
        if not mermaid_code:
            return ""
            
        try:
            # 將內容轉換為 Base64
            code_bytes = mermaid_code.encode('utf-8')
            base64_code = base64.b64encode(code_bytes).decode('utf-8')
            
            # 使用 mermaid.ink 服務
            url = f"https://mermaid.ink/png/{base64_code}"
            
            # 建立本地圖檔
            img_filename = f"chart_{os.urandom(4).hex()}.png"
            img_path = os.path.join(tmpdir, img_filename)
            
            resp = requests.get(url, timeout=20)
            if resp.status_code == 200:
                with open(img_path, "wb") as f:
                    f.write(resp.content)
                # 重要：在圖片前後加上換行，確保 Pandoc 識別為獨立段落
                return f"\n\n![Flowchart]({img_path})\n\n"
            else:
                st.error(f"Mermaid 服務回傳錯誤代碼: {resp.status_code}")
        except Exception as e:
            st.warning(f"圖表轉換過程中發生異常: {e}")
        
        return f"\n\n```mermaid\n{mermaid_code}\n```\n\n"

    return pattern.sub(download_img, md_text)

# --- UI 介面 ---
st.title("📈 專業級 Markdown 轉 Word 工具")
st.markdown("當前版本：已強化 Mermaid 匹配與圖片嵌入邏輯。")

with st.sidebar:
    st.header("⚙️ 轉換設定")
    add_toc = st.checkbox("自動生成目錄 (TOC)", value=True)
    math_support = st.checkbox("支援數學公式 ($LaTeX$)", value=True)
    ref_file = st.file_uploader("上傳參考 Word (.docx)", type=["docx"])

uploaded_files = st.file_uploader("上傳 .md 檔案", type=["md"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 開始轉換", use_container_width=True):
        zip_buffer = BytesIO()
        
        try:
            with st.spinner('正在精準匹配並轉換 Mermaid 圖表...'):
                with tempfile.TemporaryDirectory() as tmpdir:
                    ref_path = None
                    if ref_file:
                        ref_path = os.path.join(tmpdir, "template.docx")
                        with open(ref_path, "wb") as f:
                            f.write(ref_file.getbuffer())

                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        progress_bar = st.progress(0)
                        
                        for i, file in enumerate(uploaded_files):
                            raw_md = file.read().decode("utf-8")
                            
                            # 預處理：將文字語法轉為本地實體圖片
                            processed_md = process_mermaid_to_local_img(raw_md, tmpdir)
                            
                            output_filename = file.name.replace(".md", ".docx")
                            temp_docx_path = os.path.join(tmpdir, f"out_{i}.docx")
                            
                            # Pandoc 參數：確保提取媒體資源
                            args = ["--standalone", "--extract-media=."]
                            if add_toc: args.append("--toc")
                            if math_support: args.append("--mathjax")
                            if ref_path: args.append(f"--reference-doc={ref_path}")
                            
                            pypandoc.convert_text(
                                processed_md, 
                                'docx', 
                                format='md', 
                                extra_args=args, 
                                outputfile=temp_docx_path
                            )
                            
                            with open(temp_docx_path, "rb") as f:
                                zip_file.writestr(output_filename, f.read())
                            
                            progress_bar.progress((i + 1) / len(uploaded_files))

            st.success("✨ 轉換成功！請點擊下方按鈕下載。")
            st.download_button(
                label="📥 下載轉換後的 ZIP 包",
                data=zip_buffer.getvalue(),
                file_name="converted_docs.zip",
                mime="application/zip",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"轉換失敗：{str(e)}")

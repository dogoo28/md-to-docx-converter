import streamlit as st
import pypandoc
import os
import tempfile
import zipfile
import re
import base64
from io import BytesIO

# --- 設定網頁標題與風格 ---
st.set_page_config(page_title="專業 MD 轉檔器 (支援 Mermaid)", page_icon="📊", layout="wide")

# --- 函式：處理 Mermaid 語法並轉換為圖片連結 ---
def process_mermaid_blocks(md_text):
    """
    掃描 MD 文字中的 mermaid 區塊，並將其替換為 mermaid.ink 的圖片連結，
    這樣 Pandoc 才能將圖表嵌入 Word。
    """
    def generate_mermaid_url(match):
        mermaid_code = match.group(1).strip()
        # 將 Mermaid 語法進行 UTF-8 編碼
        code_bytes = mermaid_code.encode('utf-8')
        # 轉換為 Base64 字串
        base64_code = base64.b64encode(code_bytes).decode('utf-8')
        # 回傳 Markdown 圖片語法，指向 mermaid.ink 服務
        return f"\n![Mermaid Flowchart](https://mermaid.ink/png/{base64_code})\n"

    # 使用正規表示式匹配 ```mermaid ... ```
    # re.DOTALL 確保 . 可以匹配換行符號
    processed_text = re.sub(r"```mermaid\s+(.*?)\s+```", generate_mermaid_url, md_text, flags=re.DOTALL)
    return processed_text

# --- UI 介面 ---
st.title("📊 專業 Markdown 轉 Word 工具")
st.subheader("支援批次轉換、自定義範本、目錄、以及 Mermaid 流程圖")

with st.sidebar:
    st.header("⚙️ 轉換設定")
    add_toc = st.checkbox("自動生成目錄 (TOC)", value=True)
    math_support = st.checkbox("支援數學公式 ($LaTeX$)", value=True)
    process_mermaid = st.checkbox("處理 Mermaid 流程圖", value=True, help="開啟後會自動將 mermaid 語法轉為圖片")
    
    st.divider()
    
    st.subheader("🎨 自定義排版樣式")
    ref_file = st.file_uploader("上傳參考 Word 範本 (.docx)", type=["docx"], help="建議修改 Word 樣式中的『正文』與『標題』字型。")

# --- 檔案上傳區 ---
uploaded_files = st.file_uploader(
    "請選擇要轉換的 Markdown 檔案 (.md)", 
    type=["md"], 
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("🚀 開始批次轉換", use_container_width=True):
        zip_buffer = BytesIO()
        
        try:
            with st.spinner('正在分析並轉換檔案...'):
                # 建立臨時目錄來儲存樣式檔與輸出的 docx
                with tempfile.TemporaryDirectory() as tmpdir:
                    
                    # 1. 處理樣式範本
                    ref_path = None
                    if ref_file:
                        ref_path = os.path.join(tmpdir, "user_template.docx")
                        with open(ref_path, "wb") as f:
                            f.write(ref_file.getbuffer())

                    # 2. 建立 ZIP 壓縮檔
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        progress_bar = st.progress(0)
                        
                        for i, file in enumerate(uploaded_files):
                            # 讀取並解碼 Markdown 內容
                            md_content = file.read().decode("utf-8")
                            
                            # 3. 預處理 Mermaid (如果勾選)
                            if process_mermaid:
                                md_content = process_mermaid_blocks(md_content)
                            
                            output_filename = file.name.replace(".md", ".docx")
                            temp_docx_path = os.path.join(tmpdir, f"out_{i}.docx")
                            
                            # 4. 設定 Pandoc 參數
                            args = ["--standalone"]
                            if add_toc: args.append("--toc")
                            if math_support: args.append("--mathjax")
                            if ref_path: args.append(f"--reference-doc={ref_path}")
                            
                            # 5. 呼叫 Pandoc 進行轉換 (解決 RuntimeError 的實體檔案路徑寫法)
                            pypandoc.convert_text(
                                md_content, 
                                'docx', 
                                format='md', 
                                extra_args=args, 
                                outputfile=temp_docx_path
                            )
                            
                            # 6. 將轉換好的實體檔讀入 ZIP 緩衝區
                            with open(temp_docx_path, "rb") as f:
                                zip_file.writestr(output_filename, f.read())
                            
                            # 更新進度條
                            progress_bar.progress((i + 1) / len(uploaded_files))

            st.success("✨ 轉換成功！")
            
            # 7. 提供下載
            st.download_button(
                label="📥 下載轉換後的 ZIP 包",
                data=zip_buffer.getvalue(),
                file_name="converted_documents.zip",
                mime="application/zip",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"發生錯誤：{str(e)}")
            st.info("💡 提示：請檢查 Markdown 中的語法是否正確，或是 Mermaid.ink 服務是否在線。")

else:
    st.info("請上傳一個或多個 .md 檔案來開始。")

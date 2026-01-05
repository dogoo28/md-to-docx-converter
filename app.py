import streamlit as st
import pypandoc
import os
import tempfile
import zipfile
from io import BytesIO

# 設定網頁標題
st.set_page_config(page_title="專業 MD 轉檔器", page_icon="📚", layout="wide")

st.title("📚 專業 Markdown 批次轉 Word 工具")
st.markdown("本工具支援批次轉換、目錄生成及自定義 Word 樣式範本。")

# --- 側邊欄設定 ---
with st.sidebar:
    st.header("⚙️ 轉換設定")
    add_toc = st.checkbox("自動生成目錄 (TOC)", value=True)
    math_support = st.checkbox("支援數學公式 ($LaTeX$)", value=True)
    
    st.divider()
    
    st.subheader("🎨 自定義排版樣式")
    ref_file = st.file_uploader("上傳參考 Word 範本 (.docx)", type=["docx"])
    if ref_file:
        st.success("✅ 已載入樣式範本")

# --- 主界面 ---
uploaded_files = st.file_uploader(
    "請選擇要轉換的 Markdown 檔案 (.md)", 
    type=["md"], 
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("🚀 開始批次轉換", use_container_width=True):
        zip_buffer = BytesIO()
        
        try:
            with st.spinner('轉換中...'):
                # 建立臨時目錄處理實體檔案
                with tempfile.TemporaryDirectory() as tmpdir:
                    
                    # 處理參考範本路徑
                    ref_path = None
                    if ref_file:
                        ref_path = os.path.join(tmpdir, "style_template.docx")
                        with open(ref_path, "wb") as f:
                            f.write(ref_file.getbuffer())

                    # 建立 ZIP
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        progress_bar = st.progress(0)
                        
                        for i, file in enumerate(uploaded_files):
                            md_text = file.read().decode("utf-8")
                            output_filename = file.name.replace(".md", ".docx")
                            
                            # 關鍵修正：定義一個臨時輸出的實體路徑
                            temp_docx_path = os.path.join(tmpdir, f"out_{i}.docx")
                            
                            # 設定 Pandoc 參數
                            args = ["--standalone"]
                            if add_toc: args.append("--toc")
                            if math_support: args.append("--mathjax")
                            if ref_path: args.append(f"--reference-doc={ref_path}")
                            
                            # 執行轉換並指定 outputfile
                            pypandoc.convert_text(
                                md_text, 
                                'docx', 
                                format='md', 
                                extra_args=args, 
                                outputfile=temp_docx_path  # 解決 RuntimeError 的關鍵
                            )
                            
                            # 從實體檔案讀回二進位數據並存入 ZIP
                            with open(temp_docx_path, "rb") as f:
                                zip_file.writestr(output_filename, f.read())
                            
                            progress_bar.progress((i + 1) / len(uploaded_files))

            st.success("✨ 轉換完成！")
            st.download_button(
                label="📥 下載轉換後的 ZIP 包",
                data=zip_buffer.getvalue(),
                file_name="converted_docs.zip",
                mime="application/zip",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"發生錯誤：{str(e)}")

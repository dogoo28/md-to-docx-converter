import streamlit as st
import pypandoc
import os
import tempfile
import zipfile
from io import BytesIO

st.set_page_config(page_title="多檔案 MD 轉檔器", page_icon="📚")

st.title("📚 批次 Markdown 轉 Word 工具")
st.write("上傳一個或多個 MD 檔案，系統將自動打包轉換後的 Word 檔。")

# 設定區
with st.sidebar:
    st.header("轉換選項")
    add_toc = st.checkbox("產生目錄", value=True)
    ref_file = st.file_uploader("樣式範本 (選填)", type=["docx"])

# 多檔案上傳
uploaded_files = st.file_uploader("選擇多個 Markdown 檔案", type=["md"], accept_multiple_files=True)

if uploaded_files:
    if st.button(f"🚀 開始轉換 {len(uploaded_files)} 個檔案"):
        zip_buffer = BytesIO() # 建立記憶體內的 ZIP 檔案
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
            with tempfile.TemporaryDirectory() as tmpdir:
                
                # 處理樣式檔
                ref_path = None
                if ref_file:
                    ref_path = os.path.join(tmpdir, "template.docx")
                    with open(ref_path, "wb") as f:
                        f.write(ref_file.getbuffer())

                progress_bar = st.progress(0)
                
                for i, file in enumerate(uploaded_files):
                    # 讀取內容
                    md_text = file.read().decode("utf-8")
                    output_filename = file.name.replace(".md", ".docx")
                    
                    # Pandoc 參數
                    args = ["--standalone"]
                    if add_toc: args.append("--toc")
                    if ref_path: args.append(f"--reference-doc={ref_path}")
                    
                    # 轉換
                    output_data = pypandoc.convert_text(md_text, 'docx', format='md', extra_args=args)
                    
                    # 寫入 ZIP
                    zip_file.writestr(output_filename, output_data)
                    
                    # 更新進度條
                    progress_bar.progress((i + 1) / len(uploaded_files))

        st.success("全部轉換完成！")
        st.download_button(
            label="📥 下載所有 Word 檔 (.zip)",
            data=zip_buffer.getvalue(),
            file_name="converted_files.zip",
            mime="application/zip"
        )

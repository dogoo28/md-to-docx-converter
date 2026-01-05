import streamlit as st
import pypandoc
import os
import tempfile
import zipfile
from io import BytesIO

# 設定網頁標題與圖示
st.set_page_config(page_title="專業 Markdown 轉檔器", page_icon="📚", layout="wide")

st.title("📚 專業 Markdown 批次轉 Word 工具")
st.markdown("""
本工具使用 Pandoc 引擎，支援批量轉換、目錄生成及自定義 Word 樣式範本。
""")

# --- 側邊欄：設定區 ---
with st.sidebar:
    st.header("⚙️ 轉換設定")
    add_toc = st.checkbox("自動生成目錄 (TOC)", value=True)
    math_support = st.checkbox("支援數學公式 ($LaTeX$)", value=True)
    
    st.divider()
    
    st.subheader("🎨 自定義排版樣式")
    ref_file = st.file_uploader("上傳參考 Word 範本 (.docx)", type=["docx"], help="上傳一個已設定好字體與標題顏色的 Word 檔，輸出的檔案將繼承其格式。")
    if ref_file:
        st.success("✅ 已載入樣式範本")
    else:
        st.info("💡 未上傳則使用 Word 預設樣式")

# --- 主界面：檔案上傳 ---
uploaded_files = st.file_uploader(
    "請選擇要轉換的 Markdown 檔案 (.md)", 
    type=["md"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"已選擇 {len(uploaded_files)} 個檔案")
    
    if st.button("🚀 開始批次轉換", use_container_width=True):
        # 建立一個記憶體緩衝區來存放 ZIP
        zip_buffer = BytesIO()
        
        try:
            with st.spinner('轉換中，請稍候...'):
                # 建立一個臨時資料夾來處理 Pandoc 的實體檔案輸出
                with tempfile.TemporaryDirectory() as tmpdir:
                    
                    # 處理樣式範本路徑
                    ref_path = None
                    if ref_file:
                        ref_path = os.path.join(tmpdir, "template_style.docx")
                        with open(ref_path, "wb") as f:
                            f.write(ref_file.getbuffer())

                    # 建立 ZIP 檔案
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        
                        progress_bar = st.progress(0)
                        
                        for i, file in enumerate(uploaded_files):
                            # 讀取 MD 內容
                            md_text = file.read().decode("utf-8")
                            output_filename = file.name.replace(".md", ".docx")
                            
                            # 定義臨時輸出的 docx 路徑 (解決 RuntimeError 的關鍵)
                            temp_output_docx = os.path.join(tmpdir, f"temp_out_{i}.docx")
                            
                            # 設定 Pandoc 參數
                            args = ["--standalone"]
                            if add_toc:
                                args.append("--toc")
                            if math_support:
                                args.append("--mathjax")
                            if ref_path:
                                args.append(f"--reference-doc={ref_path}")
                            
                            # 執行轉換並將結果存入臨時檔案
                            pypandoc.convert_text(
                                md_text, 
                                'docx', 
                                format='md', 
                                extra_args=args, 
                                outputfile=temp_output_docx
                            )
                            
                            # 將轉換好的檔案讀入並寫進 ZIP
                            with open(temp_output_docx, "rb") as f:
                                zip_file.writestr(output_filename, f.read())
                            
                            # 更新進度條
                            progress_bar.progress((i + 1) / len(uploaded_files))

            st.success("✨ 所有檔案轉換完成！")
            
            # 提供下載按鈕
            st.download_button(
                label="📥 下載轉換後的 ZIP 包",
                data=zip_buffer.getvalue(),
                file_name="markdown_converted_docs.zip",
                mime="application/zip",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"發生錯誤：{str(e)}")
            st.warning("請確保您的 GitHub 倉庫中有 `packages.txt` 且內容為 `pandoc`。")

else:
    st.info("請先上傳 .md 檔案開始轉換。")

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

# --- 核心轉換函式：將 Markdown 轉換為 DOCX ---
def convert_md_to_docx(md_content: str, output_filename: str, options: dict, tmpdir: str) -> bytes:
    """
    核心轉換函式：將 Markdown 內容轉換為 DOCX 格式。
    
    Args:
        md_content: Markdown 文字內容
        output_filename: 輸出檔名（不含副檔名）
        options: 轉換選項 dict，包含:
            - add_toc: bool
            - math_support: bool  
            - ref_path: str (範本路徑，可為 None)
            - meta_title, meta_author, meta_date: str
        tmpdir: 臨時目錄路徑
        
    Returns:
        bytes: DOCX 檔案的二進位內容
    """
    # 處理 Mermaid 流程圖
    processed_md = process_mermaid_to_local_img(md_content, tmpdir)
    
    # 設定 Pandoc 參數
    args = ["--standalone"]
    if options.get("add_toc"): 
        args.append("--toc")
        args.append("--metadata=toc-title:目錄")
    if options.get("math_support"): 
        args.append("--mathjax")
    if options.get("ref_path"): 
        args.append(f"--reference-doc={options['ref_path']}")
    
    # Metadata 設定
    if options.get("meta_title"): 
        args.append(f"--metadata=title:{options['meta_title']}")
    if options.get("meta_author"): 
        args.append(f"--metadata=author:{options['meta_author']}")
    if options.get("meta_date"): 
        args.append(f"--metadata=date:{options['meta_date']}")
    
    # 執行轉換
    temp_docx_path = os.path.join(tmpdir, f"{output_filename}.docx")
    pypandoc.convert_text(
        processed_md, 'docx', format='md', extra_args=args, outputfile=temp_docx_path
    )
    
    # 讀取並返回二進位內容
    with open(temp_docx_path, "rb") as f:
        return f.read()

# --- 輔助函式：取得並解析 Word 樣式列表 ---
def get_docx_style_info(docx_path):
    import docx
    from docx.shared import RGBColor
    import pandas as pd
    
    # 常見樣式對照表
    STYLE_MAPPING = {
        "Normal": ("內文 (Normal)", "預設的文字樣式"),
        "Heading 1": ("標題 1 (Heading 1)", "文章主標題"),
        "Heading 2": ("標題 2 (Heading 2)", "次級標題"),
        "Heading 3": ("標題 3 (Heading 3)", "小標題"),
        "Title": ("文件標題 (Title)", "整份文件的封面標題"),
        "Subtitle": ("副標題 (Subtitle)", "文件的副標題"),
        "Author": ("作者 (Author)", "封面作者資訊"),
        "Date": ("日期 (Date)", "封面日期資訊"),
        "Abstract": ("摘要 (Abstract)", "文章摘要"),
        "Block Text": ("區塊文字 (Block Text)", "用於引用或強調的區塊"),
        "Caption": ("圖表說明 (Caption)", "圖片或表格下方的說明文字"),
        "Table Caption": ("表格標題 (Table Caption)", "表格專用的標題樣式"),
        "Image Caption": ("圖片標題 (Image Caption)", "圖片專用的標題樣式"),
        "TOC Heading": ("目錄標題 (TOC Heading)", "自動產生目錄的標題"),
        "Body Text": ("本文 (Body Text)", "一般的內文樣式"),
        "First Paragraph": ("首段 (First Paragraph)", "章節的第一段落"),
    }

    try:
        doc = docx.Document(docx_path)
        styles = []
        
        for s in doc.styles:
            if s.type == docx.enum.style.WD_STYLE_TYPE.PARAGRAPH:
                name = s.name
                # 屬性提取
                font_name = getattr(s.font, 'name', '') or '預設 (繼承)'
                
                # 字體大小 (Point)
                font_size = ''
                if s.font.size:
                    font_size = f"{s.font.size.pt} pt"
                
                # 字體顏色
                font_color = ''
                if s.font.color and s.font.color.rgb:
                    font_color = f"#{s.font.color.rgb}"
                    
                # 判斷是否為常用樣式
                if name in STYLE_MAPPING:
                    display_name, desc = STYLE_MAPPING[name]
                    category = "常用樣式"
                elif name.startswith("Heading"):
                    display_name = name
                    desc = "更深層級的標題"
                    category = "標題樣式"
                else:
                    display_name = name
                    desc = "自訂或進階樣式"
                    category = "其他樣式"
                
                styles.append({
                    "樣式名稱 (原始)": name,
                    "說明": display_name,
                    "用途": desc,
                    "字型": font_name,
                    "大小": font_size,
                    "顏色": font_color,
                    "類別": category
                })
        
        # 轉換為 DataFrame 並排序
        df = pd.DataFrame(styles)
        if not df.empty:
            cat_order = ["常用樣式", "標題樣式", "其他樣式"]
            df["類別"] = pd.Categorical(df["類別"], categories=cat_order, ordered=True)
            df = df.sort_values(["類別", "樣式名稱 (原始)"])
            
        return df
    except Exception as e:
        return pd.DataFrame([{"錯誤": f"無法讀取樣式: {str(e)}"}] )

# --- Streamlit UI 介面 ---
st.title("📈 專業級 Markdown 轉 Word 工具")
st.markdown("已優化：**支援中文流程圖**、**自動嵌入圖片**、**自定義 Word 範本**、**合併輸出**、**直接貼上文字**。")

with st.sidebar:
    st.header("⚙️ 轉換設定")
    add_toc = st.checkbox("自動生成目錄 (TOC)", value=True)
    math_support = st.checkbox("支援數學公式 ($LaTeX$)", value=True)
    merge_output = st.checkbox("合併所有檔案為單一 Word", value=False, help="將所有上傳的 MD 檔合併成一個 Docx，中間以分頁符號隔開。")
    
    st.divider()
    
    with st.expander("📝 文件屬性 (Metadata)", expanded=False):
        st.caption("適用於所有轉換模式")
        meta_title = st.text_input("文件標題", value="")
        meta_author = st.text_input("作者", value="")
        meta_date = st.text_input("日期", value="")
    
    st.subheader("🎨 樣式範本")
    template_source = st.radio("選擇範本來源", ["使用內建範本", "上傳自訂範本"])
    
    ref_file = None
    default_template_path = None
    
    if template_source == "使用內建範本":
        templates_dir = "templates"
        if not os.path.exists(templates_dir):
            os.makedirs(templates_dir)
        template_files = [f for f in os.listdir(templates_dir) if f.endswith(".docx")]
        
        if template_files:
            selected_template = st.selectbox("選擇範本", template_files)
            default_template_path = os.path.join(templates_dir, selected_template)
        else:
            st.warning("⚠️ templates 資料夾中找不到 .docx 範本，請先放入檔案。")
            st.markdown("💡 *提示：您可以執行 `pandoc --print-default-data-file reference.docx > templates/Standard.docx` 來建立預設範本*")
            
    else:
        ref_file = st.file_uploader("上傳參考 Word (.docx)", type=["docx"], help="輸出的 Word 將繼承此檔案的字型與標題樣式。")
        if ref_file:
            st.success("✅ 樣式範本已載入")

# --- 主畫面：樣式預覽 ---
target_style_source = None
source_name = ""

if template_source == "使用內建範本" and default_template_path:
    target_style_source = default_template_path
    source_name = f"內建範本 ({os.path.basename(default_template_path)})"
elif template_source == "上傳自訂範本" and ref_file:
    target_style_source = ref_file
    source_name = f"自訂範本 ({ref_file.name})"

if target_style_source:
    with st.expander(f"🎨 查看樣式詳情：{source_name}", expanded=False):
        df = get_docx_style_info(target_style_source)
        st.dataframe(
            df, 
            hide_index=True,
            column_config={
                "樣式名稱 (原始)": st.column_config.TextColumn("樣式代碼", help="Pandoc 轉換時對應的樣式名稱"),
                "說明": st.column_config.TextColumn("中文說明"),
                "用途": st.column_config.TextColumn("用途描述"),
                "字型": st.column_config.TextColumn("字型"),
                "大小": st.column_config.TextColumn("大小"),
                "顏色": st.column_config.TextColumn("顏色"),
            },
            use_container_width=True
        )

# --- 檔案處理區（標籤頁式 UI）---
st.divider()
tab_upload, tab_paste = st.tabs(["📁 上傳檔案", "📝 貼上文字"])

# --- 建立共用的轉換選項 ---
def get_conversion_options(tmpdir):
    """取得當前的轉換選項設定"""
    ref_path = None
    if default_template_path:
        ref_path = default_template_path
    elif ref_file:
        ref_path = os.path.join(tmpdir, "user_style.docx")
        with open(ref_path, "wb") as f:
            f.write(ref_file.getbuffer())
    
    return {
        "add_toc": add_toc,
        "math_support": math_support,
        "ref_path": ref_path,
        "meta_title": meta_title,
        "meta_author": meta_author,
        "meta_date": meta_date
    }

# ========================================
# 標籤 1：上傳檔案模式
# ========================================
with tab_upload:
    uploaded_files = st.file_uploader("上傳一個或多個 .md 檔案", type=["md"], accept_multiple_files=True)
    
    if uploaded_files:
        # 建立檔案名稱與物件的對照表
        file_map = {f.name: f for f in uploaded_files}
        original_filenames = [f.name for f in uploaded_files]
        
        # --- 排序設定 ---
        files_to_process = uploaded_files # 預設依上傳順序
        
        if merge_output and len(uploaded_files) > 1:
            st.subheader("🔀 調整合併順序")
            sorted_filenames = st.multiselect(
                "請依序選擇要合併的檔案 (由上而下)",
                options=original_filenames,
                default=original_filenames
            )
            
            if len(sorted_filenames) != len(uploaded_files):
                st.warning("⚠️ 請選取所有檔案以確保完整合併。")
            
            files_to_process = [file_map[name] for name in sorted_filenames]

        if st.button("🚀 開始轉換", use_container_width=True, key="upload_convert"):
            if not files_to_process:
                st.error("請至少選擇一個檔案進行轉換。")
            else:
                try:
                    with st.spinner('正在分析架構圖並轉換格式...'):
                        with tempfile.TemporaryDirectory() as tmpdir:
                            options = get_conversion_options(tmpdir)
                            
                            # 單檔且非合併模式：直接下載 .docx
                            if len(files_to_process) == 1 and not merge_output:
                                file = files_to_process[0]
                                file.seek(0)
                                raw_md = file.read().decode("utf-8")
                                output_name = file.name.replace(".md", "")
                                
                                docx_bytes = convert_md_to_docx(raw_md, output_name, options, tmpdir)
                                
                                st.success("✨ 轉換成功！")
                                st.download_button(
                                    label="📥 下載 Word 文件",
                                    data=docx_bytes,
                                    file_name=f"{output_name}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    use_container_width=True
                                )
                            
                            # 多檔或合併模式：使用 ZIP 打包
                            else:
                                zip_buffer = BytesIO()
                                
                                # 設定 Pandoc 共用參數
                                args = ["--standalone"]
                                if add_toc: args.append("--toc")
                                if math_support: args.append("--mathjax")
                                if options["ref_path"]: args.append(f"--reference-doc={options['ref_path']}")
                                if meta_title: args.append(f"--metadata=title:{meta_title}")
                                if meta_author: args.append(f"--metadata=author:{meta_author}")
                                if meta_date: args.append(f"--metadata=date:{meta_date}")
                                if add_toc: args.append("--metadata=toc-title:目錄")
                                
                                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                                    
                                    if merge_output and len(files_to_process) > 1:
                                        # --- 合併模式 ---
                                        merged_md_list = []
                                        page_break = '\n\n```{=openxml}\n<w:p><w:r><w:br w:type="page"/></w:r></w:p>\n```\n\n'
                                        
                                        for file in files_to_process:
                                            file.seek(0)
                                            raw_md = file.read().decode("utf-8")
                                            processed = process_mermaid_to_local_img(raw_md, tmpdir)
                                            merged_md_list.append(processed)
                                        
                                        merged_md = page_break.join(merged_md_list)
                                        if add_toc:
                                            merged_md = page_break + merged_md
                                        
                                        output_filename = "merged_document.docx"
                                        temp_docx_path = os.path.join(tmpdir, output_filename)
                                        
                                        pypandoc.convert_text(
                                            merged_md, 'docx', format='md', extra_args=args, outputfile=temp_docx_path
                                        )
                                        
                                        with open(temp_docx_path, "rb") as f:
                                            zip_file.writestr(output_filename, f.read())
                                    
                                    else:
                                        # --- 個別轉換模式 ---
                                        progress_bar = st.progress(0)
                                        for i, file in enumerate(files_to_process):
                                            file.seek(0)
                                            raw_md = file.read().decode("utf-8")
                                            processed_md = process_mermaid_to_local_img(raw_md, tmpdir)
                                            
                                            output_filename = file.name.replace(".md", ".docx")
                                            temp_docx_path = os.path.join(tmpdir, f"output_{i}.docx")
                                            
                                            pypandoc.convert_text(
                                                processed_md, 'docx', format='md', extra_args=args, outputfile=temp_docx_path
                                            )
                                            
                                            with open(temp_docx_path, "rb") as f:
                                                zip_file.writestr(output_filename, f.read())
                                            
                                            progress_bar.progress((i + 1) / len(files_to_process))
                                
                                st.success("✨ 轉換成功！")
                                st.download_button(
                                    label="📥 下載結果 (ZIP)",
                                    data=zip_buffer.getvalue(),
                                    file_name="converted_documents.zip",
                                    mime="application/zip",
                                    use_container_width=True
                                )
                
                except Exception as e:
                    st.error(f"轉換失敗：{str(e)}")
                    st.info("💡 提示：若出現 404，請確認您的 Mermaid 語法是否能在一般編輯器中正確顯示。")
    else:
        st.info("請上傳 .md 檔案以開始使用。")

# ========================================
# 標籤 2：貼上文字模式
# ========================================
with tab_paste:
    st.markdown("直接貼上 Markdown 內容，即可快速轉換為 Word 文件。")
    
    paste_filename = st.text_input(
        "輸出檔名", 
        value="document", 
        help="輸出的 Word 檔案名稱（不需輸入副檔名）"
    )
    
    paste_content = st.text_area(
        "Markdown 內容",
        height=400,
        placeholder="# 標題\n\n在這裡貼上您的 Markdown 內容...\n\n## 子標題\n\n- 項目一\n- 項目二",
        help="支援標準 Markdown 語法、Mermaid 流程圖、以及 LaTeX 數學公式"
    )
    
    if st.button("🚀 開始轉換", use_container_width=True, key="paste_convert"):
        if not paste_content.strip():
            st.error("請輸入 Markdown 內容。")
        elif not paste_filename.strip():
            st.error("請輸入輸出檔名。")
        else:
            try:
                with st.spinner('正在分析架構圖並轉換格式...'):
                    with tempfile.TemporaryDirectory() as tmpdir:
                        options = get_conversion_options(tmpdir)
                        
                        # 清理檔名（移除不合法字元）
                        safe_filename = re.sub(r'[<>:"/\\|?*]', '_', paste_filename.strip())
                        
                        docx_bytes = convert_md_to_docx(paste_content, safe_filename, options, tmpdir)
                        
                        st.success("✨ 轉換成功！")
                        st.download_button(
                            label="📥 下載 Word 文件",
                            data=docx_bytes,
                            file_name=f"{safe_filename}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
            
            except Exception as e:
                st.error(f"轉換失敗：{str(e)}")
                st.info("💡 提示：若出現 404，請確認您的 Mermaid 語法是否能在一般編輯器中正確顯示。")

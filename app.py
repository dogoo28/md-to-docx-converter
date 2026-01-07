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
# --- Streamlit UI 介面 ---
def load_css():
    css_path = os.path.join("assets", "custom.css")
    if os.path.exists(css_path):
        with open(css_path, "r", encoding="utf-8") as f:
            st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

load_css()

# Header Section
st.markdown("""
    <div style='text-align: center; padding: 2rem 0;'>
        <h1 style='margin-bottom: 0.5rem;'>📄 Markdown 轉 Word 專業版</h1>
        <p style='color: #666; font-size: 1.1rem;'>
            將您的 Markdown 文件轉換為格式完美的 Word 報告。<br>
            支援 <span style='background: #f0f2f6; padding: 2px 6px; border-radius: 4px; font-size: 0.9em;'>Mermaid 流程圖</span> 
            與 <span style='background: #f0f2f6; padding: 2px 6px; border-radius: 4px; font-size: 0.9em;'>LaTeX 數學公式</span>
        </p>
    </div>
""", unsafe_allow_html=True)

# --- Sidebar: Configuration ---
with st.sidebar:
    st.markdown("### ⚙️ 核心設定")
    
    with st.container():
        st.markdown('<div class="css-card" style="padding: 1rem;">', unsafe_allow_html=True)
        add_toc = st.toggle("生成目錄 (TOC)", value=True)
        math_support = st.toggle("支援數學公式", value=True)
        merge_output = st.toggle("合併輸出", value=False, help="將所有上傳的檔案合併為單一份 Word 文件")
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("### 🎨 樣式與範本")
    template_source = st.radio("來源", ["使用內建範本", "上傳自訂範本"], horizontal=True, label_visibility="collapsed")
    
    ref_file = None
    default_template_path = None
    
    if template_source == "使用內建範本":
        templates_dir = "templates"
        if not os.path.exists(templates_dir):
            os.makedirs(templates_dir)
        template_files = [f for f in os.listdir(templates_dir) if f.endswith(".docx")]
        
        if template_files:
            selected_template = st.selectbox("選擇範本樣式", template_files)
            default_template_path = os.path.join(templates_dir, selected_template)
        else:
            st.error("找不到內建範本")
    else:
        ref_file = st.file_uploader("上傳 .docx 範本", type=["docx"])

    st.divider()
    
    with st.expander("📝 文件屬性 (Metadata)"):
        meta_title = st.text_input("文件標題", placeholder="例如：年度報告")
        meta_author = st.text_input("作者", placeholder="您的姓名")
        meta_date = st.text_input("日期", placeholder="YYYY-MM-DD")

# --- Logic for Style Preview (Condensed) ---
target_style_source = default_template_path if template_source == "使用內建範本" else ref_file
if target_style_source:
    with st.expander(f"👁️ 查看範本樣式詳情 ({'內建' if default_template_path else '自訂'})"):
        df_styles = get_docx_style_info(target_style_source)
        st.dataframe(df_styles, hide_index=True, use_container_width=True)

# --- Main Content: Tabs ---
st.markdown("<div style='margin-bottom: 2rem;'></div>", unsafe_allow_html=True)
tab_upload, tab_paste = st.tabs(["📁 批量上傳轉換", "✍️ 線上編輯貼上"])

# Helper for Options
def get_conversion_options(tmpdir):
    ref_path = None
    if default_template_path:
        ref_path = default_template_path
    elif ref_file:
        ref_path = os.path.join(tmpdir, "user_style.docx")
        with open(ref_path, "wb") as f:
            f.write(ref_file.getbuffer())
    
    return {
        "add_toc": add_toc, "math_support": math_support, "ref_path": ref_path,
        "meta_title": meta_title, "meta_author": meta_author, "meta_date": meta_date
    }

# --- TAB 1: Upload ---
with tab_upload:
    st.markdown("#### 上傳 Markdown 檔案")
    uploaded_files = st.file_uploader(
        "拖放檔案至此", 
        type=["md"], 
        accept_multiple_files=True,
        label_visibility="collapsed"
    )
    
    if uploaded_files:
        col1, col2 = st.columns([2, 1])
        with col1:
             st.info(f"已選擇 {len(uploaded_files)} 個檔案")
        
        files_to_process = uploaded_files
        
        # Sort Logic if Merging
        if merge_output and len(uploaded_files) > 1:
            st.markdown("##### 🔀 調整合併順序")
            file_map = {f.name: f for f in uploaded_files}
            sorted_names = st.multiselect(
                "拖曳調整順序",
                options=[f.name for f in uploaded_files],
                default=[f.name for f in uploaded_files]
            )
            files_to_process = [file_map[n] for n in sorted_names]

        if st.button("🚀 開始轉換專案", type="primary", use_container_width=True):
            if not files_to_process:
                st.error("請至少選擇一個檔案")
            else:
                with st.spinner('正在處理文檔與圖表...'):
                    try:
                        with tempfile.TemporaryDirectory() as tmpdir:
                            options = get_conversion_options(tmpdir)
                            
                            # Single File or No Merge -> Zip or individual download?
                            # Logic: If single file, direct download. If multiple and NOT merge, Zip. If merge, direct download.
                            
                            # Case 1: Single Output (One file uploaded OR Merge all)
                            if len(files_to_process) == 1 or merge_output:
                                if merge_output and len(files_to_process) > 1:
                                    # Merge Logic
                                    merged_md_list = []
                                    page_break = '\n\n```{=openxml}\n<w:p><w:r><w:br w:type="page"/></w:r></w:p>\n```\n\n'
                                    for file in files_to_process:
                                        file.seek(0)
                                        raw = file.read().decode("utf-8")
                                        merged_md_list.append(process_mermaid_to_local_img(raw, tmpdir))
                                    final_md = page_break.join(merged_md_list)
                                    if add_toc: final_md = page_break + final_md # hack for TOC position if needed or handled by pandoc
                                    
                                    output_name = "Merged_Document"
                                else:
                                    # Single File
                                    f = files_to_process[0]
                                    f.seek(0)
                                    final_md = process_mermaid_to_local_img(f.read().decode("utf-8"), tmpdir)
                                    output_name = f.name.replace(".md", "")

                                docx_bytes = convert_md_to_docx(final_md, output_name, options, tmpdir)
                                st.success("🎉 轉換完成！")
                                st.download_button(
                                    label=f"📥 下載 {output_name}.docx",
                                    data=docx_bytes,
                                    file_name=f"{output_name}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    type="primary",
                                    use_container_width=True
                                )
                            
                            # Case 2: Multiple Files, No Merge -> ZIP
                            else:
                                zip_buffer = BytesIO()
                                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                                    # We need to run conversion for each
                                    # Reuse logic requires defining args again roughly, or calling convert_md_to_docx loop
                                    # convert_md_to_docx returns bytes, so we can use it.
                                    
                                    progress_text = st.empty()
                                    bar = st.progress(0)
                                    
                                    for i, file in enumerate(files_to_process):
                                        file.seek(0)
                                        raw = file.read().decode("utf-8")
                                        fname = file.name.replace(".md", "")
                                        docx = convert_md_to_docx(raw, fname, options, tmpdir)
                                        zf.writestr(f"{fname}.docx", docx)
                                        bar.progress((i + 1) / len(files_to_process))
                                    
                                    progress_text.text("打包完成！")
                                
                                st.success("🎉 批量轉換完成！")
                                st.download_button(
                                    label="📥 下載轉換結果 (ZIP)",
                                    data=zip_buffer.getvalue(),
                                    file_name="converted_docs.zip",
                                    mime="application/zip",
                                    type="primary",
                                    use_container_width=True
                                )

                    except Exception as e:
                        st.error(f"發生錯誤: {e}")

# --- TAB 2: Paste ---
with tab_paste:
    c1, c2 = st.columns([3, 1])
    with c1:
        st.markdown("#### 直接編輯內容")
    with c2:
        filename_input = st.text_input("輸出檔名", value="New_Document", label_visibility="collapsed")
    
    txt_input = st.text_area(
        "Markdown Input", 
        height=500,
        placeholder="# 在此貼上 Markdown...",
        label_visibility="collapsed"
    )
    
    if st.button("🚀 轉換貼上內容", key="btn_paste", type="primary", use_container_width=True):
        if not txt_input.strip():
            st.warning("⚠️ 內容不能為空")
        else:
            with st.spinner("轉換中..."):
                try:
                    with tempfile.TemporaryDirectory() as tmpdir:
                        opts = get_conversion_options(tmpdir)
                        safe_name = re.sub(r'[<>:"/\\|?*]', '_', filename_input.strip())
                        docx_bytes = convert_md_to_docx(txt_input, safe_name, opts, tmpdir)
                        
                        st.success("轉換成功！")
                        st.download_button(
                            label="📥 下載檔案",
                            data=docx_bytes,
                            file_name=f"{safe_name}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary",
                            use_container_width=True
                        )
                except Exception as e:
                    st.error(f"轉換錯誤: {e}")


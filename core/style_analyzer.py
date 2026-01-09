import docx
from docx.shared import RGBColor
import pandas as pd

# --- 輔助函式：取得並解析 Word 樣式列表 ---
def get_docx_style_info(docx_path):
    
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

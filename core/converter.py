import pypandoc
import os
import re
import base64
import requests
import streamlit as st # Need st for error display currently, in future better to return errors

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

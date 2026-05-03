"""格式转换系统 - Streamlit 主程序"""
import streamlit as st
import os
from pathlib import Path
from io import BytesIO

import converter

st.set_page_config(
    page_title="格式转换系统",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== 自定义样式 ====================
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        text-align: center;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.3rem;
    }
    .sub-header {
        text-align: center;
        color: #888;
        font-size: 1rem;
        margin-bottom: 2rem;
    }
    .feature-card {
        background: linear-gradient(135deg, #f5f7fa 0%, #e4e8ec 100%);
        border-radius: 16px;
        padding: 1.5rem;
        text-align: center;
        border: 1px solid #ddd;
        transition: transform 0.2s;
        height: 100%;
    }
    .feature-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 8px 24px rgba(0,0,0,0.1);
    }
    .feature-icon {
        font-size: 2.5rem;
        margin-bottom: 0.5rem;
    }
    .feature-title {
        font-size: 1.1rem;
        font-weight: bold;
        color: #333;
        margin-bottom: 0.3rem;
    }
    .feature-desc {
        font-size: 0.85rem;
        color: #666;
    }
    .upload-area {
        background: #f8f9fa;
        border: 2px dashed #667eea;
        border-radius: 12px;
        padding: 2rem;
        text-align: center;
    }
    .result-box {
        background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
        border-radius: 12px;
        padding: 1.5rem;
        margin-top: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# ==================== 转换类型配置 ====================

CONVERSION_TYPES = {
    "home": {
        "label": "🏠 首页",
        "icon": "🏠",
        "title": "格式转换系统",
        "desc": "支持 PDF、图像、Word 等多种格式互转"
    },
    "pdf2img": {
        "label": "📄 PDF → 图像",
        "icon": "📄",
        "title": "PDF 转图像",
        "desc": "将 PDF 页面转换为 PNG、JPG、WEBP 等图像格式",
        "input": "PDF",
        "output": "PNG / JPG / WEBP / TIFF / BMP",
        "accept": ["pdf"]
    },
    "img2pdf": {
        "label": "🖼️ 图像 → PDF",
        "icon": "🖼️",
        "title": "图像转 PDF",
        "desc": "将单张或多张图像合并为 PDF 文件",
        "input": "PNG / JPG / WEBP / BMP / TIFF / GIF",
        "output": "PDF",
        "accept": ["png", "jpg", "jpeg", "webp", "bmp", "tiff", "gif"]
    },
    "pdf2word": {
        "label": "📄 PDF → Word",
        "icon": "📄",
        "title": "PDF 转 Word",
        "desc": "将 PDF 文档转换为可编辑的 Word 文档",
        "input": "PDF",
        "output": "DOCX",
        "accept": ["pdf"]
    },
    "img2word": {
        "label": "🖼️→📝 图像转 Word",
        "icon": "🖼️",
        "title": "图像转 Word",
        "desc": "图片嵌入Word或OCR提取文字生成可编辑文档",
        "input": "PNG / JPG / WEBP / BMP / TIFF / GIF",
        "output": "DOCX",
        "accept": ["png", "jpg", "jpeg", "webp", "bmp", "tiff", "gif"]
    },
    "img2img": {
        "label": "🔄 图像格式互转",
        "icon": "🔄",
        "title": "图像格式互转",
        "desc": "PNG、JPG、WEBP、BMP、TIFF、GIF 相互转换",
        "input": "任意图像格式",
        "output": "PNG / JPG / WEBP / BMP / TIFF / GIF",
        "accept": ["png", "jpg", "jpeg", "webp", "bmp", "tiff", "gif"]
    },
    "merge_images": {
        "label": "📷 图片合并长图",
        "icon": "📷",
        "title": "图片合并长图",
        "desc": "将多张图片垂直拼接为一张长图",
        "input": "多张图像",
        "output": "PNG 长图",
        "accept": ["png", "jpg", "jpeg", "webp", "bmp", "tiff", "gif"]
    },
    "merge_pdf": {
        "label": "📑 PDF 合并",
        "icon": "📑",
        "title": "PDF 合并",
        "desc": "将多个 PDF 文件合并为一个",
        "input": "多个 PDF",
        "output": "合并后的 PDF",
        "accept": ["pdf"]
    },
    "split_pdf": {
        "label": "✂️ PDF 拆分",
        "icon": "✂️",
        "title": "PDF 拆分",
        "desc": "按页码将 PDF 拆分为多个文件",
        "input": "PDF",
        "output": "多个 PDF（ZIP 打包）",
        "accept": ["pdf"]
    },
    "extract_text": {
        "label": "📝 PDF 提取文本",
        "icon": "📝",
        "title": "PDF 提取文本",
        "desc": "从 PDF 中提取纯文本内容",
        "input": "PDF",
        "output": "TXT",
        "accept": ["pdf"]
    }
}


def render_home():
    """首页 - 展示所有功能"""
    st.markdown('<div class="main-header">🔄 格式转换系统</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">File Format Converter | 支持 PDF、图像、Word 等多种格式互转</div>', unsafe_allow_html=True)
    
    # 功能卡片网格
    features = [
        ("pdf2img", "📄→🖼️", "PDF 转图像", "PNG / JPG / WEBP / TIFF / BMP<br>支持自定义 DPI 和页码范围"),
        ("img2pdf", "🖼️→📄", "图像转 PDF", "单张或多张合并<br>保持原始分辨率"),
        ("pdf2word", "📄→📝", "PDF 转 Word", "保留排版格式<br>转换为可编辑 DOCX"),
        ("img2word", "🖼️→📝", "图像转 Word", "OCR提取文字或图片嵌入<br>生成 DOCX 文档"),
        ("img2img", "🔄", "图像格式互转", "PNG ↔ JPG ↔ WEBP<br>↔ BMP ↔ TIFF ↔ GIF"),
        ("merge_images", "📷", "图片合并长图", "多张图片垂直拼接<br>支持间距和对齐设置"),
        ("merge_pdf", "📑", "PDF 合并", "多文件一键合并<br>保持原始顺序"),
        ("split_pdf", "✂️", "PDF 拆分", "按页码拆分<br>支持自定义范围"),
        ("extract_text", "📄→📝", "PDF 提取文本", "提取纯文本内容<br>保存为 TXT 文件"),
    ]
    
    cols = st.columns(4)
    for i, (key, icon, title, desc) in enumerate(features):
        with cols[i % 4]:
            st.markdown(f"""
                <div class="feature-card">
                    <div class="feature-icon">{icon}</div>
                    <div class="feature-title">{title}</div>
                    <div class="feature-desc">{desc}</div>
                </div>
            """, unsafe_allow_html=True)
            if st.button(f"开始使用 →", key=f"btn_{key}", use_container_width=True):
                st.session_state.conversion_type = key
                st.rerun()
    
    st.divider()
    
    # 使用说明
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        #### 📋 使用步骤
        1. **选择转换类型**：点击左侧菜单或上方卡片
        2. **上传文件**：拖拽或点击上传源文件
        3. **设置参数**：选择输出格式、DPI、页码等
        4. **开始转换**：点击转换按钮
        5. **下载结果**：转换完成后点击下载链接
        """)
    with col2:
        st.markdown("""
        #### ⚡ 技术特性
        - **本地运行**：文件不上传云端，保护隐私
        - **批量处理**：支持多文件同时上传
        - **高质输出**：图像转换保持原始分辨率
        - **格式丰富**：覆盖常见办公文档格式
        - **开源部署**：代码简洁，易于二次开发
        """)
    
    st.divider()
    st.caption("💡 提示：所有转换在本地完成，文件不会上传到任何服务器。输出文件保留在 output/ 目录下。")


def render_converter(key: str):
    """渲染具体转换页面"""
    cfg = CONVERSION_TYPES[key]
    st.markdown(f"<h2>{cfg['icon']} {cfg['title']}</h2>", unsafe_allow_html=True)
    st.markdown(f"<p style='color:#666;'>{cfg['desc']}</p>", unsafe_allow_html=True)
    st.divider()
    
    if key == "pdf2img":
        render_pdf2img()
    elif key == "img2pdf":
        render_img2pdf()
    elif key == "pdf2word":
        render_pdf2word()
    elif key == "img2word":
        render_img2word()
    elif key == "img2img":
        render_img2img()
    elif key == "merge_images":
        render_merge_images()
    elif key == "merge_pdf":
        render_merge_pdf()
    elif key == "split_pdf":
        render_split_pdf()
    elif key == "extract_text":
        render_extract_text()


def render_pdf2img():
    """PDF 转图像"""
    uploaded = st.file_uploader("上传 PDF 文件", type=["pdf"], key="up_pdf2img")
    
    if uploaded:
        col1, col2 = st.columns(2)
        with col1:
            fmt = st.selectbox("输出格式", ["PNG", "JPG", "WEBP", "TIFF", "BMP"], key="fmt_pdf2img")
        with col2:
            dpi = st.slider("DPI (分辨率)", 72, 400, 200, key="dpi_pdf2img")
        
        page_range = st.text_input("页码范围（留空=全部，如: 1,3,5 或 1-5）", "", key="range_pdf2img")
        
        if st.button("🚀 开始转换", type="primary", use_container_width=True):
            with st.spinner("正在转换..."):
                pdf_bytes = uploaded.getvalue()
                page_range_val = page_range if page_range.strip() else "all"
                
                result_paths = converter.pdf_to_images(
                    pdf_bytes, fmt.lower(), dpi, page_range_val
                )
                
                if not result_paths:
                    st.error("转换失败，请检查页码范围是否正确")
                    return
                
                if len(result_paths) == 1:
                    # 单页直接下载
                    with open(result_paths[0], "rb") as f:
                        data = f.read()
                    st.download_button(
                        label=f"📥 下载 {result_paths[0].name}",
                        data=data,
                        file_name=result_paths[0].name,
                        mime=f"image/{fmt.lower()}"
                    )
                else:
                    # 多页打包zip
                    zip_path = converter.pdf_to_images_zip(
                        pdf_bytes, fmt.lower(), dpi, page_range_val
                    )
                    with open(zip_path, "rb") as f:
                        data = f.read()
                    st.download_button(
                        label=f"📥 下载 ZIP 包（{len(result_paths)} 张图像）",
                        data=data,
                        file_name=zip_path.name,
                        mime="application/zip"
                    )
                
                st.success(f"转换完成！共 {len(result_paths)} 页")


def render_img2pdf():
    """图像转 PDF"""
    uploaded = st.file_uploader(
        "上传图像文件（支持多张）",
        type=["png", "jpg", "jpeg", "webp", "bmp", "tiff", "gif"],
        accept_multiple_files=True,
        key="up_img2pdf"
    )
    
    if uploaded:
        mode = st.radio("输出模式", ["合并为单个 PDF", "各自独立输出"], key="mode_img2pdf")
        mode_val = "merge" if mode == "合并为单个 PDF" else "separate"
        
        if st.button("🚀 开始转换", type="primary", use_container_width=True):
            with st.spinner("正在转换..."):
                files = [f.getvalue() for f in uploaded]
                names = [f.name for f in uploaded]
                
                result_path = converter.images_to_pdf(files, names, mode_val)
                
                with open(result_path, "rb") as f:
                    data = f.read()
                
                mime = "application/zip" if result_path.suffix == ".zip" else "application/pdf"
                st.download_button(
                    label=f"📥 下载 {result_path.name}",
                    data=data,
                    file_name=result_path.name,
                    mime=mime
                )
                st.success("转换完成！")


def render_pdf2word():
    """PDF 转 Word"""
    uploaded = st.file_uploader("上传 PDF 文件", type=["pdf"], key="up_pdf2word")
    
    if uploaded:
        st.info("⚠️ 复杂排版（如表格、图文混排）可能无法完美还原，建议转换后手动调整。")
        
        if st.button("🚀 开始转换", type="primary", use_container_width=True):
            with st.spinner("正在转换，请稍候..."):
                pdf_bytes = uploaded.getvalue()
                result_path, mode_info = converter.pdf_to_word(pdf_bytes)
                
                with open(result_path, "rb") as f:
                    data = f.read()
                
                st.download_button(
                    label=f"📥 下载 {result_path.name}",
                    data=data,
                    file_name=result_path.name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                st.success(f"转换完成！模式: {mode_info}")


def _check_ocr_available() -> tuple[bool, str]:
    """检测 OCR 环境是否可用，返回 (是否可用, 错误信息)"""
    # 优先检查 RapidOCR（云端友好，无需 Tesseract）
    try:
        from rapidocr_onnxruntime import RapidOCR
        return True, ""
    except Exception as e:
        rapidocr_err = str(e)[:200]
    
    # 回退检查 pytesseract（本地，需要 Tesseract）
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        return True, ""
    except Exception as e:
        tesseract_err = str(e)[:200]
    
    error_msg = f"RapidOCR: {rapidocr_err}; Tesseract: {tesseract_err}"
    return False, error_msg


def render_img2word():
    """图像转 Word"""
    ocr_ok, ocr_err = _check_ocr_available()
    
    uploaded = st.file_uploader(
        "上传图像文件",
        type=["png", "jpg", "jpeg", "webp", "bmp", "tiff", "gif"],
        key="up_img2word"
    )
    
    if uploaded:
        if ocr_ok:
            st.success("✅ OCR 环境已就绪")
            mode_options = [
                "🥇 图片+文字混合（推荐：第1页原图，第2页可编辑文字）",
                "OCR文字提取（纯文字，丢失图片和格式）",
                "图片嵌入（保留原图，不可编辑文字）"
            ]
        else:
            st.error("❌ OCR 环境未就绪，当前只能使用图片嵌入模式")
            with st.expander("查看详细错误（调试用）"):
                st.code(ocr_err, language="text")
            st.info("""
            **如需 OCR 文字识别：**
            
            **本地部署**：`pip install rapidocr_onnxruntime`
            
            **Streamlit Cloud**：需在 `packages.txt` 添加 `libgl1`，并确保 `runtime.txt` 指定 Python 3.11
            """)
            mode_options = ["图片嵌入（保留原图，不可编辑文字）"]
        
        mode = st.radio("转换模式", mode_options, key="mode_img2word")
        if "混合" in mode:
            mode_val = "hybrid"
        elif "OCR" in mode:
            mode_val = "ocr"
        else:
            mode_val = "embed"
        
        if st.button("🚀 开始转换", type="primary", use_container_width=True):
            with st.spinner("正在转换..."):
                img_bytes = uploaded.getvalue()
                result_path, mode_info = converter.image_to_word(img_bytes, uploaded.name, mode_val)
                
                with open(result_path, "rb") as f:
                    data = f.read()
                
                st.download_button(
                    label=f"📥 下载 {result_path.name}",
                    data=data,
                    file_name=result_path.name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                st.success(f"转换完成！模式: {mode_info}")


def render_img2img():
    """图像格式互转"""
    uploaded = st.file_uploader(
        "上传图像文件",
        type=["png", "jpg", "jpeg", "webp", "bmp", "tiff", "gif"],
        key="up_img2img"
    )
    
    if uploaded:
        fmt = st.selectbox("输出格式", ["PNG", "JPG", "WEBP", "BMP", "TIFF", "GIF"], key="fmt_img2img")
        
        if st.button("🚀 开始转换", type="primary", use_container_width=True):
            with st.spinner("正在转换..."):
                img_bytes = uploaded.getvalue()
                result_path = converter.convert_image(img_bytes, uploaded.name, fmt.lower())
                
                with open(result_path, "rb") as f:
                    data = f.read()
                
                mime_map = {
                    "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
                    "webp": "image/webp", "bmp": "image/bmp", "tiff": "image/tiff", "gif": "image/gif"
                }
                st.download_button(
                    label=f"📥 下载 {result_path.name}",
                    data=data,
                    file_name=result_path.name,
                    mime=mime_map.get(fmt.lower(), "application/octet-stream")
                )
                st.success("转换完成！")


def render_merge_images():
    """图片合并长图"""
    uploaded = st.file_uploader(
        "上传图片（支持多张，按上传顺序拼接）",
        type=["png", "jpg", "jpeg", "webp", "bmp", "tiff", "gif"],
        accept_multiple_files=True,
        key="up_merge_img"
    )
    
    if uploaded and len(uploaded) >= 2:
        col1, col2, col3 = st.columns(3)
        with col1:
            spacing = st.slider("图片间距", 0, 50, 10, key="img_space")
        with col2:
            align = st.selectbox("对齐方式", ["居中", "左对齐", "右对齐"], key="img_align")
        with col3:
            max_w = st.number_input("最大宽度", min_value=100, max_value=3000, value=1200, step=100, key="img_maxw")
        
        align_map = {"居中": "center", "左对齐": "left", "右对齐": "right"}
        
        if st.button("🚀 开始拼接", type="primary", use_container_width=True):
            with st.spinner("正在拼接..."):
                files = [f.getvalue() for f in uploaded]
                names = [f.name for f in uploaded]
                result_path = converter.merge_images_vertical(
                    files, names,
                    spacing=spacing,
                    align=align_map[align],
                    max_width=max_w
                )
                
                with open(result_path, "rb") as f:
                    data = f.read()
                
                st.download_button(
                    label=f"📥 下载 {result_path.name}",
                    data=data,
                    file_name=result_path.name,
                    mime="image/png"
                )
                st.success(f"拼接完成！共 {len(uploaded)} 张图片 → {result_path.name}")
    elif uploaded:
        st.warning("请至少上传 2 张图片进行拼接")


def render_merge_pdf():
    """PDF 合并"""
    uploaded = st.file_uploader(
        "上传 PDF 文件（按上传顺序合并）",
        type=["pdf"],
        accept_multiple_files=True,
        key="up_merge"
    )
    
    if uploaded and len(uploaded) >= 2:
        st.info(f"已上传 {len(uploaded)} 个文件，将按上传顺序合并。")
        
        if st.button("🚀 开始合并", type="primary", use_container_width=True):
            with st.spinner("正在合并..."):
                files = [f.getvalue() for f in uploaded]
                names = [f.name for f in uploaded]
                result_path = converter.merge_pdfs(files, names)
                
                with open(result_path, "rb") as f:
                    data = f.read()
                
                st.download_button(
                    label=f"📥 下载 {result_path.name}",
                    data=data,
                    file_name=result_path.name,
                    mime="application/pdf"
                )
                st.success(f"合并完成！共 {len(uploaded)} 个文件")
    elif uploaded:
        st.warning("请至少上传 2 个 PDF 文件进行合并")


def render_split_pdf():
    """PDF 拆分"""
    uploaded = st.file_uploader("上传 PDF 文件", type=["pdf"], key="up_split")
    
    if uploaded:
        mode = st.radio("拆分方式", ["每页独立", "按范围拆分"], key="mode_split")
        
        if mode == "按范围拆分":
            ranges = st.text_input("拆分范围（如: 1-3,4-6,7-9）", "1-3,4-6", key="range_split")
        else:
            ranges = ""
        
        if st.button("🚀 开始拆分", type="primary", use_container_width=True):
            with st.spinner("正在拆分..."):
                pdf_bytes = uploaded.getvalue()
                mode_val = "single" if mode == "每页独立" else "range"
                result_path = converter.split_pdf(pdf_bytes, mode_val, ranges)
                
                with open(result_path, "rb") as f:
                    data = f.read()
                
                mime = "application/zip" if result_path.suffix == ".zip" else "application/pdf"
                st.download_button(
                    label=f"📥 下载 {result_path.name}",
                    data=data,
                    file_name=result_path.name,
                    mime=mime
                )
                st.success("拆分完成！")


def render_extract_text():
    """PDF 提取文本"""
    uploaded = st.file_uploader("上传 PDF 文件", type=["pdf"], key="up_extract")
    
    if uploaded:
        if st.button("🚀 开始提取", type="primary", use_container_width=True):
            with st.spinner("正在提取文本..."):
                pdf_bytes = uploaded.getvalue()
                result_path = converter.extract_pdf_text(pdf_bytes)
                
                with open(result_path, "rb") as f:
                    data = f.read()
                
                # 同时显示文本预览
                text_content = result_path.read_text(encoding="utf-8")
                preview = text_content[:2000] + ("..." if len(text_content) > 2000 else "")
                
                st.markdown("**文本预览：**")
                st.text_area("", preview, height=300, key="text_preview")
                
                st.download_button(
                    label=f"📥 下载 {result_path.name}",
                    data=data,
                    file_name=result_path.name,
                    mime="text/plain"
                )
                st.success("提取完成！")


# ==================== 主程序 ====================

def main():
    # 清理旧文件
    converter.cleanup_old_files()
    
    with st.sidebar:
        st.markdown("## 🔄 格式转换系统")
        st.markdown("---")
        
        # 菜单选择
        options = [(k, v["label"]) for k, v in CONVERSION_TYPES.items()]
        labels = [v["label"] for v in CONVERSION_TYPES.values()]
        
        # 从 session_state 恢复选择
        if "conversion_type" not in st.session_state:
            st.session_state.conversion_type = "home"
        
        selected_label = st.radio(
            "选择转换功能",
            labels,
            index=labels.index(CONVERSION_TYPES[st.session_state.conversion_type]["label"])
        )
        
        # 根据 label 找到 key
        for k, v in CONVERSION_TYPES.items():
            if v["label"] == selected_label:
                st.session_state.conversion_type = k
                break
        
        st.markdown("---")
        st.caption("v1.0 · 本地处理 · 不上传云端")
    
    # 渲染页面
    key = st.session_state.conversion_type
    if key == "home":
        render_home()
    else:
        render_converter(key)


if __name__ == "__main__":
    main()

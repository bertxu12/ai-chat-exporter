import streamlit as st
import os
from datetime import datetime
from io import BytesIO
import re

# PDF 生成
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import cm

# Word 生成
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# Excel 生成
import pandas as pd
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


# ==================== 对话解析器 ====================
class DialogParser:
    """智能解析对话内容"""

    @staticmethod
    def parse(text):
        """解析对话文本为结构化数据"""
        lines = text.strip().split('\n')
        messages = []
        current_role = None
        current_content = []

        role_keywords = {
            'user': ['用户', '我', 'User', 'Me', 'Human', '人类'],
            'assistant': ['AI', 'Grok', 'Claude', 'ChatGPT', 'GPT', 'Assistant', '助手', 'Bot', '机器人']
        }

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # 检测角色切换
            role_detected = None
            content = line

            if '：' in line or ':' in line:
                separator = '：' if '：' in line else ':'
                parts = line.split(separator, 1)
                role_part = parts[0].strip()

                # 判断角色
                for role_type, keywords in role_keywords.items():
                    if any(kw in role_part for kw in keywords):
                        role_detected = role_type
                        content = parts[1].strip() if len(parts) > 1 else ''
                        break

            # 如果检测到新角色，保存之前的内容
            if role_detected and role_detected != current_role:
                if current_role and current_content:
                    messages.append({
                        'role': current_role,
                        'content': '\n'.join(current_content)
                    })
                current_role = role_detected
                current_content = [content] if content else []
            else:
                # 继续累积当前角色的内容
                if current_role:
                    current_content.append(line)
                else:
                    # 如果还没检测到角色，默认为用户
                    current_role = 'user'
                    current_content = [line]

        # 保存最后一条消息
        if current_role and current_content:
            messages.append({
                'role': current_role,
                'content': '\n'.join(current_content)
            })

        return messages


# ==================== PDF 导出器（优化版）====================
class PDFExporter:
    @staticmethod
    def register_fonts():
        """注册中文字体"""
        try:
            font_paths = [
                ('C:/Windows/Fonts/msyh.ttc', 'YaHei'),  # 微软雅黑
                ('C:/Windows/Fonts/simhei.ttf', 'SimHei'),  # 黑体
                ('C:/Windows/Fonts/simsun.ttc', 'SimSun'),  # 宋体
                ('/System/Library/Fonts/PingFang.ttc', 'PingFang'),  # macOS
                ('/usr/share/fonts/truetype/wqy/wqy-microhei.ttc', 'WQY'),  # Linux
            ]

            for path, name in font_paths:
                if os.path.exists(path):
                    pdfmetrics.registerFont(TTFont(name, path))
                    return name
        except:
            pass
        return 'Helvetica'

    @staticmethod
    def create_styles(font_name):
        """创建优化的PDF样式"""

        return {
            'title': ParagraphStyle(
                'CustomTitle',
                fontName=font_name,
                fontSize=24,
                alignment=TA_CENTER,
                spaceAfter=30,
                textColor=colors.HexColor('#1a1a1a'),
                leading=30
            ),
            'subtitle': ParagraphStyle(
                'Subtitle',
                fontName=font_name,
                fontSize=11,
                alignment=TA_CENTER,
                spaceAfter=20,
                textColor=colors.HexColor('#666666')
            ),
            'user_role': ParagraphStyle(
                'UserRole',
                fontName=font_name,
                fontSize=12,
                textColor=colors.HexColor('#2563eb'),
                spaceAfter=8,
                leftIndent=0,
                spaceBefore=10
            ),
            'assistant_role': ParagraphStyle(
                'AssistantRole',
                fontName=font_name,
                fontSize=12,
                textColor=colors.HexColor('#16a34a'),
                spaceAfter=8,
                leftIndent=0,
                spaceBefore=10
            ),
            'content': ParagraphStyle(
                'Content',
                fontName=font_name,
                fontSize=11,
                leading=20,
                spaceAfter=18,
                leftIndent=15,
                rightIndent=15,
                textColor=colors.HexColor('#2d3748'),
                alignment=TA_JUSTIFY
            )
        }

    @staticmethod
    def generate(messages, title="AI对话记录"):
        """生成优化的PDF"""
        buffer = BytesIO()
        font_name = PDFExporter.register_fonts()
        styles = PDFExporter.create_styles(font_name)

        pdf = SimpleDocTemplate(
            buffer,
            pagesize=A4,
            leftMargin=2.5 * cm,
            rightMargin=2.5 * cm,
            topMargin=3 * cm,
            bottomMargin=2.5 * cm
        )

        elements = []

        # 标题页
        elements.append(Spacer(1, 1 * cm))
        elements.append(Paragraph(title, styles['title']))
        elements.append(Spacer(1, 0.3 * cm))
        elements.append(Paragraph(
            f"导出时间：{datetime.now().strftime('%Y年%m月%d日 %H:%M')}",
            styles['subtitle']
        ))
        elements.append(Paragraph(
            f"共 {len(messages)} 轮对话",
            styles['subtitle']
        ))
        elements.append(Spacer(1, 1.5 * cm))

        # 对话内容
        for i, msg in enumerate(messages, 1):
            role_emoji = "👤" if msg['role'] == 'user' else "🤖"
            role_name = "用户" if msg['role'] == 'user' else "AI助手"
            role_style = styles['user_role'] if msg['role'] == 'user' else styles['assistant_role']

            # 角色标签
            elements.append(Paragraph(
                f"<b>{role_emoji} {role_name} (第{i}轮)</b>",
                role_style
            ))

            # 内容 - 保留换行
            content = msg['content'].replace('\n', '<br/>')
            content = content.replace('<', '&lt;').replace('>', '&gt;')
            elements.append(Paragraph(content, styles['content']))

            # 添加分隔线（除了最后一条）
            if i < len(messages):
                elements.append(Spacer(1, 0.2 * cm))

        pdf.build(elements)
        buffer.seek(0)
        return buffer


# ==================== Word 导出器（优化版）====================
class WordExporter:
    @staticmethod
    def set_cell_border(cell, **kwargs):
        """设置单元格边框"""
        tc = cell._element
        tcPr = tc.get_or_add_tcPr()

        # 创建边框
        tcBorders = tcPr.first_child_found_in("w:tcBorders")
        if tcBorders is None:
            tcBorders = parse_xml(r'<w:tcBorders %s/>' % nsdecls('w'))
            tcPr.append(tcBorders)

    @staticmethod
    def generate(messages, title="AI对话记录"):
        """生成优化的Word文档"""
        doc = Document()

        # 设置中文字体
        doc.styles['Normal'].font.name = '微软雅黑'
        doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

        # 标题
        heading = doc.add_heading(title, 0)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading_run = heading.runs[0]
        heading_run.font.size = Pt(24)
        heading_run.font.color.rgb = RGBColor(26, 26, 26)
        heading_run.font.name = '微软雅黑'

        # 导出信息
        info_para = doc.add_paragraph()
        info_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        info_run = info_para.add_run(
            f"导出时间：{datetime.now().strftime('%Y年%m月%d日 %H:%M')} | 共 {len(messages)} 轮对话"
        )
        info_run.font.size = Pt(10)
        info_run.font.color.rgb = RGBColor(102, 102, 102)

        doc.add_paragraph()  # 空行

        # 对话内容
        for i, msg in enumerate(messages, 1):
            role_emoji = "👤" if msg['role'] == 'user' else "🤖"
            role_name = "用户" if msg['role'] == 'user' else "AI助手"
            role_color = RGBColor(37, 99, 235) if msg['role'] == 'user' else RGBColor(22, 163, 74)
            bg_color = RGBColor(239, 246, 255) if msg['role'] == 'user' else RGBColor(240, 253, 244)

            # 角色标签
            role_para = doc.add_paragraph()
            role_run = role_para.add_run(f"{role_emoji} {role_name} (第{i}轮)")
            role_run.bold = True
            role_run.font.size = Pt(12)
            role_run.font.color.rgb = role_color
            role_run.font.name = '微软雅黑'

            # 内容段落 - 添加背景色效果
            content_para = doc.add_paragraph()
            content_para.paragraph_format.left_indent = Cm(0.5)
            content_para.paragraph_format.right_indent = Cm(0.5)
            content_para.paragraph_format.space_after = Pt(15)
            content_para.paragraph_format.line_spacing = 1.5

            content_run = content_para.add_run(msg['content'])
            content_run.font.size = Pt(11)
            content_run.font.name = '微软雅黑'
            content_run.font.color.rgb = RGBColor(45, 55, 72)

            # 添加分隔线
            if i < len(messages):
                separator = doc.add_paragraph()
                separator.paragraph_format.space_before = Pt(5)
                separator.paragraph_format.space_after = Pt(5)

        # 保存到内存
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer


# ==================== Excel 导出器（优化版）====================
class ExcelExporter:
    @staticmethod
    def generate(messages, title="AI对话记录"):
        """生成优化的Excel表格"""
        # 准备数据
        data = []
        for i, msg in enumerate(messages, 1):
            role = "👤 用户" if msg['role'] == 'user' else "🤖 AI助手"
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')
            data.append({
                '序号': i,
                '角色': role,
                '内容': msg['content'],
                '字数': len(msg['content']),
                '时间': timestamp
            })

        # 创建DataFrame
        df = pd.DataFrame(data)

        # 保存到内存
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='对话记录', index=False)

            # 美化样式
            workbook = writer.book
            worksheet = writer.sheets['对话记录']

            # 设置列宽
            worksheet.column_dimensions['A'].width = 8
            worksheet.column_dimensions['B'].width = 15
            worksheet.column_dimensions['C'].width = 100
            worksheet.column_dimensions['D'].width = 10
            worksheet.column_dimensions['E'].width = 18

            # 标题行样式
            header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
            header_font = Font(bold=True, color='FFFFFF', size=12, name='微软雅黑')

            thin_border = Border(
                left=Side(style='thin', color='CCCCCC'),
                right=Side(style='thin', color='CCCCCC'),
                top=Side(style='thin', color='CCCCCC'),
                bottom=Side(style='thin', color='CCCCCC')
            )

            # 应用标题样式
            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border

            # 内容样式
            for row_idx, row in enumerate(worksheet.iter_rows(min_row=2, max_row=worksheet.max_row), start=2):
                for cell in row:
                    cell.alignment = Alignment(vertical='top', wrap_text=True)
                    cell.border = thin_border
                    cell.font = Font(name='微软雅黑', size=10)

                # 角色列颜色标记
                role_cell = row[1]
                if '用户' in str(role_cell.value):
                    role_cell.font = Font(color='2563EB', bold=True, name='微软雅黑', size=10)
                    role_cell.fill = PatternFill(start_color='EFF6FF', end_color='EFF6FF', fill_type='solid')
                else:
                    role_cell.font = Font(color='16A34A', bold=True, name='微软雅黑', size=10)
                    role_cell.fill = PatternFill(start_color='F0FDF4', end_color='F0FDF4', fill_type='solid')

                # 序号列居中
                row[0].alignment = Alignment(horizontal='center', vertical='center')

                # 字数列居中
                row[3].alignment = Alignment(horizontal='center', vertical='center')

                # 设置行高
                worksheet.row_dimensions[row_idx].height = max(20, len(str(row[2].value)) / 50 * 15)

            # 冻结首行
            worksheet.freeze_panes = 'A2'

        buffer.seek(0)
        return buffer


# ==================== Streamlit 界面 ====================
def main():
    st.set_page_config(
        page_title="AI对话导出工具",
        page_icon="💬",
        layout="wide"
    )

    # 自定义CSS
    st.markdown("""
        <style>
        .main > div {
            padding-top: 2rem;
        }
        .stButton > button {
            width: 100%;
            border-radius: 8px;
            height: 50px;
            font-weight: 600;
        }
        .success-box {
            padding: 1rem;
            border-radius: 8px;
            background-color: #d1fae5;
            border-left: 4px solid #10b981;
            margin: 1rem 0;
        }
        </style>
    """, unsafe_allow_html=True)

    st.title("💬 AI对话多格式导出工具")
    st.markdown("### 📱 支持导出 PDF / Word / Excel 三种格式")
    st.markdown("---")

    # 侧边栏设置
    with st.sidebar:
        st.header("⚙️ 导出设置")
        export_title = st.text_input("📝 对话标题", value="AI对话记录")

        st.markdown("#### 选择导出格式")
        export_pdf = st.checkbox("📄 PDF格式", value=True)
        export_word = st.checkbox("📝 Word格式", value=False)
        export_excel = st.checkbox("📊 Excel格式", value=False)

        st.markdown("---")
        st.markdown("### 📖 使用说明")
        st.markdown("""
        1. **复制对话内容**
           - 从AI聊天界面复制完整对话

        2. **粘贴到文本框**
           - 保持原有格式即可

        3. **选择导出格式**
           - PDF：精美排版，适合阅读
           - Word：方便编辑修改
           - Excel：数据分析统计

        4. **点击导出下载**
           - 自动生成文件下载

        **✅ 支持平台：**
        - ChatGPT / Claude
        - Grok / Gemini
        - 文心一言 / 通义千问
        - 其他所有对话AI
        """)

        st.markdown("---")
        st.markdown("### 💡 提示")
        st.info("支持中英文混合对话，自动识别角色身份")

    # 主界面
    col1, col2 = st.columns([2, 1])

    with col1:
        st.subheader("📝 粘贴对话内容")
        dialog_text = st.text_area(
            "将完整对话内容粘贴到这里（支持多轮对话）",
            height=450,
            placeholder="""示例格式：

用户：你好，请帮我写一个Python脚本

AI：你好！我很乐意帮你写Python脚本。请告诉我你想要实现什么功能？

用户：我想做一个文件批量重命名的工具

AI：好的，我来帮你写一个文件批量重命名脚本...

（继续复制更多对话内容）""",
            key="dialog_input"
        )

    with col2:
        st.subheader("📊 解析预览")
        if dialog_text:
            messages = DialogParser.parse(dialog_text)

            # 统计信息
            col_a, col_b = st.columns(2)
            with col_a:
                st.metric("🔄 对话轮次", len(messages))
            with col_b:
                total_chars = sum(len(msg['content']) for msg in messages)
                st.metric("📝 总字数", f"{total_chars:,}")

            # 显示前3条预览
            with st.expander("🔍 查看解析结果（前3条）", expanded=True):
                for idx, msg in enumerate(messages[:3], 1):
                    role_emoji = "👤" if msg['role'] == 'user' else "🤖"
                    role_name = "用户" if msg['role'] == 'user' else "AI助手"

                    st.markdown(f"**{role_emoji} {role_name} (第{idx}轮)**")
                    preview_text = msg['content'][:150]
                    if len(msg['content']) > 150:
                        preview_text += "..."
                    st.text(preview_text)
                    st.markdown("---")

                if len(messages) > 3:
                    st.info(f"还有 {len(messages) - 3} 轮对话未显示...")
        else:
            st.info("👈 请在左侧粘贴对话内容")

    st.markdown("---")

    # 导出按钮区域
    if dialog_text and (export_pdf or export_word or export_excel):
        st.subheader("📥 导出文件")

        messages = DialogParser.parse(dialog_text)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        col1, col2, col3, col4 = st.columns([1, 1, 1, 1])

        with col1:
            if export_pdf and st.button("📄 生成PDF", use_container_width=True, type="primary"):
                with st.spinner("🔄 正在生成PDF文件..."):
                    pdf_buffer = PDFExporter.generate(messages, export_title)
                    st.success("✅ PDF生成成功！")
                    st.download_button(
                        label="⬇️ 下载PDF文件",
                        data=pdf_buffer,
                        file_name=f"{export_title}_{timestamp}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )

        with col2:
            if export_word and st.button("📝 生成Word", use_container_width=True, type="primary"):
                with st.spinner("🔄 正在生成Word文档..."):
                    word_buffer = WordExporter.generate(messages, export_title)
                    st.success("✅ Word生成成功！")
                    st.download_button(
                        label="⬇️ 下载Word文档",
                        data=word_buffer,
                        file_name=f"{export_title}_{timestamp}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )

        with col3:
            if export_excel and st.button("📊 生成Excel", use_container_width=True, type="primary"):
                with st.spinner("🔄 正在生成Excel表格..."):
                    excel_buffer = ExcelExporter.generate(messages, export_title)
                    st.success("✅ Excel生成成功！")
                    st.download_button(
                        label="⬇️ 下载Excel表格",
                        data=excel_buffer,
                        file_name=f"{export_title}_{timestamp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

        with col4:
            if st.button("🔄 清空重置", use_container_width=True):
                st.rerun()

    elif dialog_text:
        st.warning("⚠️ 请在左侧至少选择一种导出格式")

    # 底部信息
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: #6b7280; padding: 20px;'>
            <p style='margin: 5px 0;'>💡 <b>提示</b>：支持从任何AI对话平台复制粘贴 | 自动识别对话结构 | 智能排版美化</p>
            <p style='margin: 5px 0;'>⭐ 如果觉得好用，欢迎分享给朋友！</p>
        </div>
        """,
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
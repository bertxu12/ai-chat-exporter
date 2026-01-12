import streamlit as st
import re
from datetime import datetime
from io import BytesIO

# PDF
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import cm

# Word
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# Excel
import pandas as pd


def clean_special_chars(text: str, aggressive: bool = False, preserve_code: bool = True) -> str:
    """
    修复版清理函数 - 避免删除代码块，保留JS/HTML等内容
    """
    if not text:
        return text

    # 先提取并保护多行代码块（用特殊标记包裹，防止后续正则干扰）
    code_blocks = []

    def replace_code_block(match):
        code_content = match.group(0)[3:-3].strip()  # 去掉```和语言标识，保留纯内容
        code_blocks.append(code_content)
        return f"< preserved_code_{len(code_blocks) - 1} >"  # 临时占位符

    if preserve_code:
        text = re.sub(r'```[\s\S]*?```', replace_code_block, text)
    else:
        # 如果不保留，直接删除（原有行为）
        text = re.sub(r'```[\s\S]*?```', '', text)

    # 2. 行内代码 → 只保留内容（但保留在上下文中）
    text = re.sub(r'`([^`]+)`', r'\1', text)

    # 3. 链接 → 只保留显示文字
    text = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', text)
    text = re.sub(r'!\[([^\]]*)\]\([^)]*\)', r'\1', text)

    # 4. 标题符号
    text = re.sub(r'^#{1,6}\s+', '', text, flags=re.MULTILINE)

    # 5. 列表符号 → 转成缩进（不直接删除内容）
    text = re.sub(r'^\s*([-*+•◦➤]|(\d+[.)]))\s+', '  • ', text, flags=re.MULTILINE)

    # 6. 清理强调、删除线 - 更安全版本
    for mark in [r'\*{1,3}', r'_{1,2}', r'~~']:
        pattern = rf'({mark})(.+?)({mark})(?!\S)'
        text = re.sub(pattern, r'\2', text, flags=re.DOTALL)

    # 清理孤立标记
    text = re.sub(r'\*{2,3}|_{2,3}|~~|\*\*', '', text)

    # 7. 移除表情符号和常见装饰字符
    text = re.sub(
        r'[\U0001F300-\U0001F9FF\U0001FA00-\U0001FAFF'
        r'\U00002700-\U000027BF\U00002600-\U000026FF'
        r'\U0001F000-\U0001FFFF]+', '', text)
    text = re.sub(r'[★☆♡♥♦♠♣●○◆◇■□▲△▼▽◀▶※♪♫✓✔✕✖]', '', text)

    # 8. 激进模式（只做最必要的过滤）
    if aggressive:
        text = re.sub(
            r'[^\u4e00-\u9fffa-zA-Z0-9\s'
            r'\u3000-\u303F\uFF00-\uFFEF'  # 中文标点 + 全角
            r'。，、；：？！…—～·（）【】《》“”‘’\'\"-.,;:!?()%+*/=&@#$^]',
            '', text)

    # 9. 收尾规范化
    text = re.sub(r'\n\s*\n\s*\n+', '\n\n', text)  # 压缩多空行
    text = re.sub(r'[ \t]{2,}', ' ', text)  # 多空格 → 单空格
    text = re.sub(r'\s+([，。、；：？！）】》”])', r'\1', text)  # 中文标点前去空格

    # 最后，放回保护的代码块（可选：添加换行和缩进以保持可读性）
    for i, code in enumerate(code_blocks):
        formatted_code = '\n'.join('    ' + line for line in code.split('\n'))  # 添加缩进，模拟代码格式
        text = text.replace(f"< preserved_code_{i} >", f"\n[代码块]\n{formatted_code}\n[/代码块]")

    return text.strip()


def parse_dialog(text: str) -> list:
    """对话解析 - 保持不变"""
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    messages = []
    current_role = None
    current_content = []

    user_keywords = {'用户', '我', 'user', 'me', 'human'}
    ai_keywords = {'ai', 'grok', 'claude', 'chatgpt', 'gpt', 'assistant', '助手', 'bot'}

    for line in lines:
        role = None
        content = line

        if '：' in line or ':' in line:
            sep = '：' if '：' in line else ':'
            parts = line.split(sep, 1)
            role_part = parts[0].strip().lower()
            content_part = parts[1].strip() if len(parts) > 1 else ''

            if any(k in role_part for k in user_keywords):
                role = 'user'
                content = content_part
            elif any(k in role_part for k in ai_keywords):
                role = 'assistant'
                content = content_part

        if role:
            if current_role and current_content:
                messages.append({
                    'role': current_role,
                    'content': '\n'.join(current_content).strip()
                })
            current_role = role
            current_content = [content] if content else []
        else:
            if current_role:
                current_content.append(line)
            else:
                current_role = 'user'
                current_content = [line]

    if current_role and current_content:
        messages.append({
            'role': current_role,
            'content': '\n'.join(current_content).strip()
        })

    return messages


def parse_markdown_tables(content):
    """
    解析 content 中的所有 Markdown 表格，返回列表：每个元素是 (pre_text, header, rows, post_text)
    如果没有表格，返回 [(content, None, None, '')]
    """
    parts = []
    last_end = 0
    table_pattern = r'(^\|.*?\|\n)+(\|[-|: ]+\|\n)((\|.*?\|\n)+)'
    for match in re.finditer(table_pattern, content, re.MULTILINE):
        pre_text = content[last_end:match.start()].strip()
        table_text = match.group(0)
        post_text = ''  # 后续文本在下一次处理

        lines = table_text.strip().split('\n')
        header = [cell.strip() for cell in lines[0][1:-1].split('|')]  # 去掉首尾 |
        separator = lines[1]
        rows = []
        for line in lines[2:]:
            row = [cell.strip() for cell in line[1:-1].split('|')]
            rows.append(row)

        parts.append((pre_text, header, rows, post_text))
        last_end = match.end()

    post_text = content[last_end:].strip()
    if parts:
        # 添加最后的后置文本到最后一个 part
        if post_text:
            pre, h, r, _ = parts[-1]
            parts[-1] = (pre, h, r, post_text)
    else:
        parts.append((content, None, None, ''))

    return parts


def generate_pdf(messages, title):
    buffer = BytesIO()

    try:
        pdfmetrics.registerFont(TTFont('YaHei', 'C:/Windows/Fonts/msyh.ttc'))
        font = 'YaHei'
    except:
        font = 'Helvetica'

    styles = {
        'title': ParagraphStyle('title', fontName=font, fontSize=22, alignment=TA_CENTER, spaceAfter=24),
        'meta': ParagraphStyle('meta', fontName=font, fontSize=10, alignment=TA_CENTER, spaceAfter=12,
                               textColor=colors.grey),
        'role': ParagraphStyle('role', fontName=font, fontSize=14, spaceBefore=12, spaceAfter=6),
        'content': ParagraphStyle('content', fontName=font, fontSize=11, leading=16, alignment=TA_JUSTIFY,
                                  spaceAfter=16)
    }

    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=2.5 * cm, rightMargin=2.5 * cm,
                            topMargin=3 * cm, bottomMargin=2.5 * cm)

    elements = [
        Paragraph(title, styles['title']),
        Paragraph(f"导出时间：{datetime.now().strftime('%Y年%m月%d日 %H:%M')}", styles['meta']),
        Paragraph(f"消息数量：{len(messages)} 条", styles['meta']),
        Spacer(1, 1.2 * cm)
    ]

    for i, msg in enumerate(messages, 1):
        role_text = "用户" if msg['role'] == 'user' else "AI助手"
        color = colors.HexColor('#2563eb') if msg['role'] == 'user' else colors.HexColor('#16a34a')

        elements.append(Paragraph(
            f"<font color='{color.hexval()}'><b>{role_text}（第 {i} 轮）</b></font>",
            styles['role']
        ))

        # 解析表格和代码
        parts = parse_markdown_tables(msg['content'])
        for pre_text, header, rows, post_text in parts:
            if pre_text:
                # 处理代码块
                content_parts = re.split(r'\[代码块\](.*?)\[/代码块\]', pre_text, flags=re.DOTALL)
                for part in content_parts:
                    if part.strip():
                        if re.match(r'^\s*\n', part):  # 代码
                            code_style = ParagraphStyle('code', fontName='Courier', fontSize=10, leading=12,
                                                        backColor=colors.lightgrey, spaceAfter=12)
                            elements.append(Paragraph(part.replace('\n', '<br/>'), code_style))
                        else:
                            elements.append(Paragraph(part.replace('\n', '<br/>'), styles['content']))

            if header and rows:
                data = [header] + rows
                table = Table(data)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, -1), font),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                elements.append(table)
                elements.append(Spacer(1, 0.5 * cm))

            if post_text:
                # 处理代码块
                content_parts = re.split(r'\[代码块\](.*?)\[/代码块\]', post_text, flags=re.DOTALL)
                for part in content_parts:
                    if part.strip():
                        if re.match(r'^\s*\n', part):  # 代码
                            code_style = ParagraphStyle('code', fontName='Courier', fontSize=10, leading=12,
                                                        backColor=colors.lightgrey, spaceAfter=12)
                            elements.append(Paragraph(part.replace('\n', '<br/>'), code_style))
                        else:
                            elements.append(Paragraph(part.replace('\n', '<br/>'), styles['content']))

        if i < len(messages):
            elements.append(Spacer(1, 0.5 * cm))

    doc.build(elements)
    buffer.seek(0)
    return buffer


def generate_word(messages, title):
    doc = Document()

    # 标题
    p = doc.add_paragraph(title)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.runs[0]
    run.font.size = Pt(22)
    run.bold = True

    # 元信息
    p = doc.add_paragraph(
        f"导出时间：{datetime.now().strftime('%Y年%m月%d日 %H:%M')}\n"
        f"消息数量：{len(messages)} 条"
    )
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.runs[0]
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(128, 128, 128)

    doc.add_paragraph()

    for i, msg in enumerate(messages, 1):
        p = doc.add_paragraph()
        role_text = f"用户（第 {i} 轮）" if msg['role'] == 'user' else f"AI助手（第 {i} 轮）"
        run = p.add_run(role_text)
        run.bold = True
        run.font.size = Pt(14)

        if msg['role'] == 'user':
            run.font.color.rgb = RGBColor(37, 99, 235)
        else:
            run.font.color.rgb = RGBColor(22, 163, 74)

        # 解析表格和代码
        parts = parse_markdown_tables(msg['content'])
        for pre_text, header, rows, post_text in parts:
            if pre_text:
                content_parts = re.split(r'\[代码块\](.*?)\[/代码块\]', pre_text, flags=re.DOTALL)
                for part in content_parts:
                    if part.strip():
                        p = doc.add_paragraph(part)
                        if re.match(r'^\s*\n', part):  # 代码
                            for run in p.runs:
                                run.font.name = 'Courier New'
                                run.font.size = Pt(10)
                            p.paragraph_format.left_indent = Pt(20)  # 缩进

            if header and rows:
                table = doc.add_table(rows=len(rows) + 1, cols=len(header))
                table.style = 'Table Grid'  # 使用网格样式
                hdr_cells = table.rows[0].cells
                for j, h in enumerate(header):
                    hdr_cells[j].text = h
                    hdr_cells[j].paragraphs[0].runs[0].bold = True
                    hdr_cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                for row_idx, row_data in enumerate(rows, 1):
                    row_cells = table.rows[row_idx].cells
                    for j, cell_text in enumerate(row_data):
                        row_cells[j].text = cell_text
                        row_cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                # 调整列宽（可选）
                for column in table.columns:
                    column.width = Inches(2.0)  # 根据需要调整

            if post_text:
                content_parts = re.split(r'\[代码块\](.*?)\[/代码块\]', post_text, flags=re.DOTALL)
                for part in content_parts:
                    if part.strip():
                        p = doc.add_paragraph(part)
                        if re.match(r'^\s*\n', part):  # 代码
                            for run in p.runs:
                                run.font.name = 'Courier New'
                                run.font.size = Pt(10)
                            p.paragraph_format.left_indent = Pt(20)  # 缩进

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def generate_excel(messages, title):
    data = []
    for i, msg in enumerate(messages, 1):
        data.append({
            '轮次': i,
            '角色': '用户' if msg['role'] == 'user' else 'AI助手',
            '内容': msg['content'],
            '字数': len(msg['content'])
        })

    df = pd.DataFrame(data)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='对话记录', index=False)
        ws = writer.sheets['对话记录']
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 80
        ws.column_dimensions['D'].width = 10

    buffer.seek(0)
    return buffer


def main():
    st.set_page_config(page_title="AI对话导出工具", page_icon="💬", layout="wide")

    if "original_text" not in st.session_state:
        st.session_state.original_text = ""
    if "current_text" not in st.session_state:
        st.session_state.current_text = ""
    if "cleaned_once" not in st.session_state:
        st.session_state.cleaned_once = False

    st.title("AI 对话导出工具")
    st.caption("修复版 - 重点解决「清理后中文消失」问题")

    with st.sidebar:
        st.header("导出设置")
        title = st.text_input("文档标题", "AI对话记录")

        st.divider()
        st.subheader("导出格式")
        export_word = st.checkbox("Word (.docx)", True)
        export_pdf = st.checkbox("PDF (.pdf)", False)
        export_excel = st.checkbox("Excel (.xlsx)", False)

        st.divider()
        st.subheader("文本清理")
        auto_clean = st.checkbox("导出时自动清理", True)
        aggressive = st.checkbox("激进模式（最大程度去干扰）", False)
        preserve_code = st.checkbox("保留代码块（如JS/HTML）", True)  # 新增

    col1, col2 = st.columns([2, 1])

    with col1:
        st.subheader("对话内容")

        raw_text = st.text_area(
            "请粘贴完整对话...",
            value=st.session_state.current_text,
            height=500
        )

        st.session_state.current_text = raw_text

        btn_col1, btn_col2, btn_col3 = st.columns([1, 1, 2])

        with btn_col1:
            if st.button("🧹 清理文本", type="primary"):
                if raw_text.strip():
                    if not st.session_state.cleaned_once:
                        st.session_state.original_text = raw_text
                    cleaned = clean_special_chars(raw_text, aggressive=aggressive, preserve_code=preserve_code)
                    st.session_state.current_text = cleaned
                    st.session_state.cleaned_once = True
                    st.success("清理完成")
                    st.rerun()

        with btn_col2:
            if st.button("↩️ 恢复原始"):
                if st.session_state.original_text:
                    st.session_state.current_text = st.session_state.original_text
                    st.session_state.cleaned_once = False
                    st.rerun()

        with btn_col3:
            if st.button("🗑️ 清空"):
                for key in ["current_text", "original_text", "cleaned_once"]:
                    if key in st.session_state:
                        del st.session_state[key]
                st.rerun()

    with col2:
        st.subheader("统计信息")
        if st.session_state.current_text.strip():
            messages = parse_dialog(st.session_state.current_text)
            if messages:
                st.metric("消息数量", len(messages))
                st.metric("总字符数", f"{sum(len(m['content']) for m in messages):,}")

    # 导出部分
    if st.session_state.current_text.strip():
        messages = parse_dialog(st.session_state.current_text)

        final_messages = messages
        if auto_clean:
            final_messages = []
            for m in messages:
                cleaned = clean_special_chars(m['content'], aggressive=aggressive, preserve_code=preserve_code)
                final_messages.append({'role': m['role'], 'content': cleaned})

        if messages and (export_pdf or export_word or export_excel):
            st.divider()
            st.subheader("导出")

            cols = st.columns(3)

            with cols[0]:
                if export_word and st.button("生成 Word"):
                    buf = generate_word(final_messages, title)
                    st.download_button(
                        "⬇️ 下载 Word", buf,
                        f"{title}_{datetime.now():%Y%m%d_%H%M}.docx",
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            with cols[1]:
                if export_pdf and st.button("生成 PDF"):
                    buf = generate_pdf(final_messages, title)
                    st.download_button(
                        "⬇️ 下载 PDF", buf,
                        f"{title}_{datetime.now():%Y%m%d_%H%M}.pdf",
                        "application/pdf"
                    )

            with cols[2]:
                if export_excel and st.button("生成 Excel"):
                    buf = generate_excel(final_messages, title)
                    st.download_button(
                        "⬇️ 下载 Excel", buf,
                        f"{title}_{datetime.now():%Y%m%d_%H%M}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )


if __name__ == "__main__":
    main()
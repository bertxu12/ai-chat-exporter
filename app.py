import streamlit as st
import re
from datetime import datetime
from io import BytesIO

# PDF
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import cm

# Word
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# Excel
import pandas as pd


def clean_special_chars(text: str, aggressive: bool = False) -> str:
    """
    修复版清理函数 - 重点避免中文被误删
    """
    if not text:
        return text

    # 1. 删除整块代码（最先处理，避免干扰后续匹配）
    text = re.sub(r'```[\s\S]*?```', '', text)

    # 2. 行内代码 → 只保留内容
    text = re.sub(r'`([^`]+)`', r'\1', text)

    # 3. 链接 → 只保留显示文字
    text = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', text)
    text = re.sub(r'!\[([^\]]*)\]\([^)]*\)', r'\1', text)

    # 4. 标题符号
    text = re.sub(r'^#{1,6}\s+', '', text, flags=re.MULTILINE)

    # 5. 列表符号 → 转成缩进（不直接删除内容）
    text = re.sub(r'^\s*([-*+•◦➤]|(\d+[.)]))\s+', '  • ', text, flags=re.MULTILINE)

    # 6. 清理强调、删除线 - 更安全版本（避免跨行大吞噬）
    # 只匹配较短的合理范围内的标记
    for mark in [r'\*{1,3}', r'_{1,2}', r'~~']:
        pattern = rf'({mark})(.+?)({mark})(?!\S)'
        text = re.sub(pattern, r'\2', text, flags=re.DOTALL)

    # 清理孤立标记（没有内容的 ** __ ~~ 等）
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
            r'\u3000-\u303F\uFF00-\uFFEF'           # 中文标点 + 全角
            r'。，、；：？！…—～·（）【】《》“”‘’\'\"-.,;:!?()%+*/=&@#$^]',
            '', text)

    # 9. 收尾规范化
    text = re.sub(r'\n\s*\n\s*\n+', '\n\n', text)     # 压缩多空行
    text = re.sub(r'[ \t]{2,}', ' ', text)           # 多空格 → 单空格
    text = re.sub(r'\s+([，。、；：？！）】》”])', r'\1', text)  # 中文标点前去空格

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

    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=2.5*cm, rightMargin=2.5*cm,
                            topMargin=3*cm, bottomMargin=2.5*cm)

    elements = [
        Paragraph(title, styles['title']),
        Paragraph(f"导出时间：{datetime.now().strftime('%Y年%m月%d日 %H:%M')}", styles['meta']),
        Paragraph(f"消息数量：{len(messages)} 条", styles['meta']),
        Spacer(1, 1.2*cm)
    ]

    for i, msg in enumerate(messages, 1):
        role_text = "用户" if msg['role'] == 'user' else "AI助手"
        color = colors.HexColor('#2563eb') if msg['role'] == 'user' else colors.HexColor('#16a34a')

        elements.append(Paragraph(
            f"<font color='{color.hexval()}'><b>{role_text}（第 {i} 轮）</b></font>",
            styles['role']
        ))

        content = msg['content'].replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        content = content.replace('\n', '<br/>')

        elements.append(Paragraph(content, styles['content']))
        if i < len(messages):
            elements.append(Spacer(1, 0.5*cm))

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

        doc.add_paragraph(msg['content'])

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

    col1, col2 = st.columns([2, 1])

    with col1:
        st.subheader("对话内容")

        raw_text = st.text_area(
            "请粘贴完整对话...",
            value=st.session_state.current_text,
            height=500
        )

        st.session_state.current_text = raw_text

        btn_col1, btn_col2, btn_col3 = st.columns([1,1,2])

        with btn_col1:
            if st.button("🧹 清理文本", type="primary"):
                if raw_text.strip():
                    if not st.session_state.cleaned_once:
                        st.session_state.original_text = raw_text
                    cleaned = clean_special_chars(raw_text, aggressive=aggressive)
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
                cleaned = clean_special_chars(m['content'], aggressive=aggressive)
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
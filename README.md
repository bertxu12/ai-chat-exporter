# 💬 AI对话多格式导出工具

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](你的应用链接)

一键导出AI对话记录为PDF、Word、Excel格式的在线工具。

## ✨ 功能特点

- 📄 **PDF导出** - 精美排版，适合分享阅读
- 📝 **Word导出** - 方便二次编辑修改  
- 📊 **Excel导出** - 结构化数据，便于分析统计
- 🎨 **智能解析** - 自动识别用户和AI角色
- 🌐 **支持多平台** - ChatGPT、Claude、Grok、文心一言等所有AI平台
- 🎯 **格式美化** - 自动排版，颜色标记，阅读体验佳

## 🚀 在线使用

👉 [立即使用](你的streamlit链接)

## 📸 预览截图

![应用截图](https://via.placeholder.com/800x400?text=App+Screenshot)

## 💻 本地运行

### 前置要求
- Python 3.8+
- pip

### 安装步骤

1. **克隆仓库**
```bash
git clone https://github.com/你的用户名/ai-dialog-exporter.git
cd ai-dialog-exporter
```

2. **安装依赖**
```bash
pip install -r requirements.txt
```

3. **启动应用**
```bash
streamlit run app.py
```

4. **访问应用**
浏览器自动打开 `http://localhost:8501`

## 📖 使用方法

### 第一步：复制对话
从任何AI聊天平台（ChatGPT、Claude等）复制完整对话内容

### 第二步：粘贴内容
将对话内容粘贴到应用的文本框中

### 第三步：选择格式
勾选需要的导出格式：
- PDF - 适合阅读分享
- Word - 适合编辑修改
- Excel - 适合数据分析

### 第四步：导出下载
点击生成按钮，等待几秒即可下载文件

## 🎨 导出格式说明

### PDF格式特点
- 使用微软雅黑字体，清晰易读
- 蓝色标记用户，绿色标记AI
- 自动添加对话轮次编号
- 合理的行距和页边距

### Word格式特点
- 统一使用微软雅黑字体
- 浅色背景区分不同角色
- 1.5倍行距，阅读舒适
- 方便后续编辑和批注

### Excel格式特点
- 结构化表格展示
- 自动统计字数和时间
- 冻结首行方便滚动
- 颜色标记角色身份

## 🛠️ 技术栈

- **框架**: Streamlit - 快速构建Web应用
- **PDF生成**: ReportLab - 专业PDF处理库
- **Word生成**: python-docx - Office文档处理
- **Excel生成**: openpyxl + pandas - 数据表格处理
- **部署**: Streamlit Cloud - 免费托管服务

## 📦 项目结构

```
ai-dialog-exporter/
├── app.py              # 主应用程序
├── requirements.txt    # Python依赖
├── README.md          # 项目说明
├── .gitignore         # Git忽略配置
└── LICENSE            # 开源协议
```

## 🤝 贡献指南

欢迎提交Issue和Pull Request！

### 开发流程
1. Fork本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交改动 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 提交Pull Request

## 📝 更新日志

### v1.0.0 (2024-01-10)
- ✅ 初始版本发布
- ✅ 支持PDF、Word、Excel三种格式导出
- ✅ 智能对话解析功能
- ✅ 美化格式排版
- ✅ 在线部署上线

## 🐛 已知问题

- [ ] 超长对话可能导致导出较慢
- [ ] 特殊字符可能影响解析准确性

## 🔮 未来计划

- [ ] 支持Markdown格式导出
- [ ] 添加对话统计分析功能
- [ ] 支持批量导出多个对话
- [ ] 添加自定义主题配置
- [ ] 支持文件拖拽上传

## ❓ 常见问题

### Q: 支持哪些AI平台？
A: 支持所有对话式AI平台，包括但不限于ChatGPT、Claude、Grok、文心一言、通义千问等。

### Q: 导出的文件可以编辑吗？
A: Word和Excel格式可以直接编辑，PDF需要使用专业工具。

### Q: 有字数限制吗？
A: 理论上无限制，但超长对话建议分段导出。

### Q: 数据安全吗？
A: 所有处理都在浏览器本地完成，不上传任何数据到服务器。

## 📄 开源协议

本项目采用 [MIT License](LICENSE) 开源协议。

## 💖 致谢

感谢以下开源项目：
- [Streamlit](https://streamlit.io/)
- [ReportLab](https://www.reportlab.com/)
- [python-docx](https://python-docx.readthedocs.io/)
- [openpyxl](https://openpyxl.readthedocs.io/)

## 📧 联系方式

- 提交Issue: [GitHub Issues](https://github.com/你的用户名/ai-dialog-exporter/issues)
- 项目主页: [GitHub Repository](https://github.com/你的用户名/ai-dialog-exporter)

---

⭐ 如果这个项目对你有帮助，请给个Star支持一下！
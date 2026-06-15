# 📄 PDF Quote Extractor - Streamlit App

从PDF报价单中自动提取信息，生成结构化的Excel表格。

**🌐 在线试用：** [Streamlit Cloud](https://quote-extractor.streamlit.app)

---

## ✨ 功能特性

✅ **自动表格识别** - 智能检测PDF中的报价表格  
✅ **列自动映射** - 自动识别和映射表头列  
✅ **多格式输出** - Excel / JSON / CSV  
✅ **数据审查** - 提取前查看和验证数据  
✅ **专业格式** - 蓝色表头、冻结窗格、格式优化  
✅ **完全本地** - 无需外部服务，数据不上传服务器  

---

## 🚀 快速开始

### 本地运行

```bash
# 1. 克隆仓库
git clone https://github.com/yourusername/pdf-quote-extractor.git
cd pdf-quote-extractor

# 2. 安装依赖
pip install -r requirements_streamlit.txt

# 3. 运行应用
streamlit run streamlit_app.py
```

应用将在 `http://localhost:8501` 打开

### 部署到Streamlit Cloud

1. **推送到GitHub**
   ```bash
   git add .
   git commit -m "Initial commit"
   git push origin main
   ```

2. **在Streamlit Cloud中部署**
   - 访问 [share.streamlit.io](https://share.streamlit.io)
   - 连接你的GitHub账户
   - 选择此仓库
   - 选择主文件：`streamlit_app.py`
   - 点击"Deploy"

3. **获取分享链接**
   - 应用部署后，你将获得一个公共URL
   - 任何人都可以通过这个URL使用你的应用

---

## 📋 支持的列

```
Numéro
Libellé / Désignation
Unité
Quantité
PU EXW excluding tax $
Montant Total EXW excluding tax $
```

---

## 🎯 使用场景

### 1. 采购人员
快速处理供应商报价单，生成标准化的Excel表格用于对比分析。

### 2. 数据分析
导出JSON或CSV数据，用于进一步分析和集成到其他系统。

### 3. 自动化工作流
与Make、n8n等无代码工具集成，实现自动化报价处理。

---

## 📊 工作流程

```
1. 上传PDF报价单
    ↓
2. 应用自动识别表格
    ↓
3. 审查提取的数据
    ↓
4. 下载 Excel / JSON / CSV
```

---

## 🔒 隐私和安全

✅ **数据完全本地处理** - 所有PDF处理都在浏览器或应用服务器本地进行  
✅ **不保存任何文件** - 上传的PDF不会被保存  
✅ **无需API密钥** - 不依赖任何外部服务  

---

## 🛠️ 技术栈

- **Streamlit** - Web框架
- **pdfplumber** - PDF表格提取
- **openpyxl** - Excel生成
- **pandas** - 数据处理

---

## 📁 项目结构

```
pdf-quote-extractor/
├── streamlit_app.py              # 主应用文件
├── requirements_streamlit.txt    # Python依赖
├── .streamlit/
│   └── config.toml              # Streamlit配置
├── .gitignore                    # Git忽略文件
├── README.md                     # 项目说明（本文件）
└── LICENSE                       # 开源许可证（可选）
```

---

## 📝 使用示例

### 基础使用
1. 打开应用
2. 点击"选择PDF文件"上传报价单
3. 点击"🚀 开始提取"
4. 在"📊 审查数据"标签页检查结果
5. 在"💾 导出结果"标签页下载Excel/JSON

### 高级用法
1. 导出JSON格式
2. 用文本编辑器修改数据（如果有错误）
3. 用其他工具（如Python脚本）进一步处理

---

## ⚙️ 配置

### 修改支持的列

编辑 `streamlit_app.py` 中的 `TARGET_COLUMNS` 变量：

```python
TARGET_COLUMNS = [
    'Numéro',
    'Libellé / Désignation',
    'Unité',
    'Quantité',
    'PU EXW excluding tax $',
    'Montant Total EXW excluding tax $'
    # 添加更多列...
]
```

### 自定义主题

编辑 `.streamlit/config.toml`：

```toml
[theme]
primaryColor = "#366092"        # 改变主色
backgroundColor = "#FFFFFF"     # 改变背景
```

---

## 🐛 故障排除

### 问题1：找不到表格
- 确保PDF包含可提取的文本（非扫描版）
- 检查表格是否使用标准格式

### 问题2：列识别不正确
- 导出为JSON查看原始数据
- 检查PDF的列标题是否匹配规则
- 可能需要调整 `_map_columns()` 方法的关键字

### 问题3：上传文件时出错
- 检查文件大小是否超过限制（默认200MB）
- 确保是有效的PDF格式

---

## 📚 相关资源

- [Streamlit文档](https://docs.streamlit.io)
- [pdfplumber文档](https://github.com/jsvine/pdfplumber)
- [openpyxl文档](https://openpyxl.readthedocs.io)

---

## 🤝 贡献

欢迎提交问题(Issue)和拉取请求(PR)！

---

## 📄 许可证

MIT License - 详见 LICENSE 文件

---

## 👤 作者

**Zélie**  
- 采购类目经理 @ NGE/SADE
- 13年采购经验
- 热爱数据驱动和自动化

---

## 💬 问题反馈

如有问题，请：
1. 检查 [FAQ](#常见问题)
2. 提交 [Issue](https://github.com/yourusername/pdf-quote-extractor/issues)
3. 发送邮件（如适用）

---

## 🎉 更新日志

### v1.0 (2026-06-15)
- ✨ 初始发布
- 🎯 支持基本的表格识别和提取
- 📊 Excel/JSON/CSV导出
- 🎨 专业的Web界面

---

**祝使用愉快！** ✨


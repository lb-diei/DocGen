# 📄 Document Generator (DocGen)

**Professional Document Template System & Document Formatter**

> *Your document automation assistant - Create and format perfect documents in seconds!*

DocGen helps you:
- 📝 **Generate** professional documents from Word templates
- 🎨 **Format** existing documents to standard styles (GB/T 9704-2012)
- 🖥️ **GUI Interface** for easy template selection and style customization

---

## 🛠️ Core Features

### 1. Document Generator (模板生成)
Generate documents from pre-designed Word templates with variables.

### 2. Document Formatter (文档格式化)
Format any Word document to standard styles:
- Auto-detect heading levels
- Apply consistent fonts and sizes
- Set margins and line spacing

### 3. GUI Frontend (图形界面)
Easy-to-use interface for:
- Selecting templates
- Customizing fonts, sizes, alignment
- Previewing and applying styles

---

## 🚀 Quick Start

### Option 1: GUI Interface (Recommended)
```bash
python doc_gen_gui.py
```
Opens an interactive window for easy document generation and formatting.

### Option 2: Command Line - Document Generation
```bash
# Generate from template
python document_generator.py notice -o 我的通知.docx

# With custom variables
python document_generator.py notice -o 年终通知.docx \
  -v title="关于2025年度工作总结的通知" \
  -v author="人力资源部"
```

### Option 3: Command Line - Document Formatting
```bash
# Format a Word document to standard style
python doc_formatter.py 原始文档.docx -o 格式化后.docx

# With custom style
python doc_formatter.py 原始.docx -o 输出.docx --preview
```

### Output Example
```
Available templates:
  - notice        (政府公文 - 通知)
  - request       (政府公文 - 请示报告)
  - notification  (企业公文 - 内部通知)
  - meeting       (企业公文 - 会议纪要)
  - report        (企业公文 - 工作报告)
  - invitation    (企业公文 - 邀请函)
  - contract      (法律文书 - 合同)
  - authorization (法律文书 - 授权委托书)
  - paper         (学术论文 - 论文)
  - thesis        (学术论文 - 毕业论文)
  - custom        (自定义模板)
```

---

## 📖 What It Does

| Input | Output |
|-------|---------|
| Word Template (.docx) | Professional DOCX Document |
| Your Variables | Automatically Filled Content |
| Template Name | Ready-to-Use File |


## 📁 Project Structure

```
DocGen/
├── document_generator.py     # Main CLI - generate DOCX from templates
├── doc_formatter.py          # Format documents to standard styles
├── doc_gen_gui.py            # 🖥️ Graphical interface (recommended)
├── templates/               # 📂 11 Word templates ready to use
│   ├── government/         # 🏛️ Government documents
│   │   ├── notice.docx     # Official notice template
│   │   └── request.docx    # Request document template
│   ├── enterprise/         # 🏢 Business documents
│   │   ├── notification.docx
│   │   ├── meeting.docx
│   │   ├── report.docx
│   │   └── invitation.docx
│   ├── legal/              # ⚖️ Legal documents
│   │   ├── contract.docx
│   │   └── authorization.docx
│   ├── academic/           # 🎓 Academic documents
│   │   ├── paper.docx
│   │   └── thesis.docx
│   └── custom/             # ✏️ Your custom template
│       └── custom.docx     # Blank template - edit freely!
├── README.md
└── requirements.txt
```

> 📝 **Note**: All templates are `.docx` files - edit them directly in Microsoft Word!


## 📋 Available Templates (11 Templates)

### 🏛️ Government Documents (政府公文)
| Template | Chinese Name | Description |
|----------|--------------|-------------|
| `notice` | 通知模板 | Official notice for internal/external communication |
| `request` | 请示报告模板 | Formal request document for superior approval |

### 🏢 Enterprise Documents (企业公文)
| Template | Chinese Name | Description |
|----------|--------------|-------------|
| `notification` | 内部通知模板 | Internal company announcements |
| `meeting` | 会议纪要模板 | Meeting minutes with action items |
| `report` | 工作报告模板 | Work reports with data tables |
| `invitation` | 邀请函模板 | Event invitations with schedules |

### ⚖️ Legal Documents (法律文书)
| Template | Chinese Name | Description |
|----------|--------------|-------------|
| `contract` | 合同模板 | Standard contract with dual signatures |
| `authorization` | 授权委托书模板 | Authorization letters with agent details |

### 🎓 Academic Documents (学术论文)
| Template | Chinese Name | Description |
|----------|--------------|-------------|
| `paper` | 学术论文模板 | Academic paper with abstract & references |
| `thesis` | 毕业论文模板 | Graduate thesis with all required sections |

### ✏️ Custom Template (自定义模板)
| Template | Chinese Name | Description |
|----------|--------------|-------------|
| `custom` | 用户自定义模板 | **Blank template for your own designs!** |

---

### 🌟 Why Use DocGen?

| Feature | Benefit |
|---------|---------|
| **Word-Based** | Edit templates directly in Microsoft Word |
| **11 Templates** | Cover 90% of common document needs |
| **Variable System** | Automate repetitive content |
| **MIT License** | Free for personal & commercial use |
| **Open Source** | Customize & extend as you need |


## How to Create Templates

Create a Word document (.docx) in `templates/` directory with placeholders:

```
{{title}}     - Document title
{{author}}    - Author name
{{date}}      - Date
{{content}}   - Main content
{{variable}}  - Any custom variable
```

### Example Placeholders

| Placeholder | Example Value |
|-------------|---------------|
| {{title}} | 关于开展2026年度工作的通知 |
| {{author}} | 人力资源部 |
| {{date}} | 2026-02-10 |
| {{content}} | 具体内容描述... |
| {{meeting_date}} | 2026年1月15日 |
| {{location}} | 会议室A |


## Usage Examples

### List All Templates

```bash
python document_generator.py -l
```

Output:
```
Available templates:
  - notice
  - request
  - notification
  - meeting
  - report
  - invitation
  - contract
  - authorization
  - paper
  - thesis
  - custom
```

### Generate with Defaults

```bash
python document_generator.py notice -o output.docx
```

### Generate with Custom Variables

```bash
python document_generator.py notice \
  -o report.docx \
  -v title="年度通知" \
  -v author="人事部"
```


## ⚡ Command Options

| Option | Description | Example |
|--------|-------------|---------|
| `template` | Template name (without .docx) | `notice` |
| `-o, --output` | Output filename (default: output.docx) | `-o mydoc.docx` |
| `-l, --list` | List all available templates | `-l` |
| `-v, --variable` | Add variable (key=value) | `-v title="通知"` |

### Examples

```bash
# List all templates
python document_generator.py -l

# Generate with defaults
python document_generator.py notice -o output.docx

# Generate with custom title and author
python document_generator.py notice \
  -o 通知.docx \
  -v title="年终通知" \
  -v author="人事部"

# Use custom template
python document_generator.py custom -o 我的文档.docx
```


## ✏️ Custom Template - Create Your Own!

DocGen includes a **special blank template** designed for you to create custom documents!

### 🎯 How to Use the Custom Template

```bash
# Generate a document from the blank custom template
python document_generator.py custom -o my_document.docx

# Or customize with your own variables
python document_generator.py custom -o mydoc.docx \
  -v title="My Custom Title" \
  -v author="My Name" \
  -v date="2026-02-10"
```

### 🚀 Create Your Own Template

1. **Open** `templates/custom/custom.docx` in Microsoft Word
2. **Edit** the document layout and styling as needed
3. **Add** placeholders like `{{title}}`, `{{author}}`, `{{date}}`
4. **Save** to a new file (e.g., `templates/enterprise/my_report.docx`)
5. **Use** your new template:

```bash
python document_generator.py my_report -o output.docx
```

### 📝 Placeholder Variables

| Placeholder | Example |
|-------------|---------|
| `{{title}}` | 年度工作总结 |
| `{{author}}` | 张三 |
| `{{date}}` | 2026-02-10 |
| `{{content}}` | 主体内容... |
| Any custom name! | `{{department}}`, `{{project}}` |

> 💡 **Pro Tip**: Use meaningful variable names like `{{meeting_date}}`, `{{deadline}}`, `{{budget}}` to make your templates self-documenting!


## 📦 Requirements

| Package | Version |
|---------|---------|
| python-docx | >=1.1.0 |

## 📄 License

MIT License - Free to use and modify

## 👤 Author

Created with Claude Code

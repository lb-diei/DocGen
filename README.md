# 📄 DocFormatter

**Document Template System - Generate Professional DOCX Documents from Word Templates**



## What It Does

| Input | Output |
|-------|---------|
| Word Template (.docx) | DOCX Document |
| User Variables | Formatted Content |
| Template Name | Ready-to-Use File |


## Quick Start

```bash
# Install dependencies
pip install python-docx

# List available templates
python document_generator.py -l

# Generate a document
python document_generator.py notice -o my_document.docx

# Generate with custom variables
python document_generator.py notice -o my_doc.docx -v title="My Title" -v author="John"
```


## Template Structure

```
lb03/
├── document_generator.py    # Main program
├── templates/              # Template directory
│   ├── government/        # Government documents
│   │   └── notice.docx  # Notice template
│   ├── enterprise/      # Business documents
│   │   └── notification.docx
│   ├── legal/           # Legal documents
│   │   └── contract.docx
│   ├── academic/        # Academic papers
│   │   └── paper.docx
│   └── custom/         # Custom templates
│       └── custom.docx
├── README.md
└── requirements.txt
```


## Available Templates

### Government Documents (政府公文)
| Template | Description | Language |
|----------|-------------|----------|
| notice | 正式通知模板 | 中文 |
| request | 请示报告模板 | 中文 |

### Enterprise Documents (企业公文)
| Template | Description | Language |
|----------|-------------|----------|
| notification | 内部通知模板 | 中文 |
| meeting | 会议纪要模板 | 中文 |
| report | 工作报告模板 | 中文 |
| invitation | 邀请函模板 | 中文 |

### Legal Documents (法律文书)
| Template | Description | Language |
|----------|-------------|----------|
| contract | 合同模板 | 中文 |
| authorization | 授权委托书模板 | 中文 |

### Academic Documents (学术论文)
| Template | Description | Language |
|----------|-------------|----------|
| paper | 学术论文格式 | 中文 |
| thesis | 毕业论文模板 | 中文 |

### Custom (自定义模板)
| Template | Description | Language |
|----------|-------------|----------|
| custom | 用户自定义模板 | 中文 |


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


## Command Options

| Option | Description |
|--------|-------------|
| template | Template name (without .docx) |
| -o, --output | Output filename (default: output.docx) |
| -l, --list | List available templates |
| -v, --variable | Add variable (key=value) |


## Add Custom Template

### Use Built-in Custom Template

1. Edit `templates/custom/custom.docx`
2. Replace placeholders with your own content
3. Use the template:

```bash
python document_generator.py custom -o mydoc.docx
```

### Create New Template

1. Create a new Word document (.docx)
2. Add placeholders where needed (e.g., {{title}}, {{author}}, {{date}})
3. Save in appropriate folder (templates/government/, templates/enterprise/, etc.)
4. Use the template:

```bash
python document_generator.py your_template_name -o output.docx
```


## Requirements

| Package | Version |
|---------|---------|
| python-docx | >=1.1.0 |


## License

MIT License - Free to use and modify


## Author

Created with Claude Code

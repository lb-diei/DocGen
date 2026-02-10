# 📄 Document Template System

**Generate Professional DOCX Documents from Customizable Templates**



## What It Does

| Input | Output |
|-------|---------|
| Markdown Template | DOCX Document |
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
│   │   └── notice.md      # Notice template
│   ├── enterprise/        # Business documents
│   │   └── notification.md
│   ├── legal/             # Legal documents
│   │   └── contract.md
│   └── academic/          # Academic papers
│       └── paper.md
├── README.md
└── requirements.txt
```


## Available Templates

### Government Documents (政府公文)
| Template | Description | 语言 |
|----------|-------------|------|
| notice | 正式通知模板 | 中文 |
| official | 正式公文模板 | 中文 |
| request | 请示报告模板 | 中文 |

### Enterprise Documents (企业公文)
| Template | Description | 语言 |
|----------|-------------|------|
| notification | 内部通知模板 | 中文 |
| meeting | 会议纪要模板 | 中文 |
| report | 工作报告模板 | 中文 |
| invitation | 邀请函模板 | 中文 |

### Legal Documents (法律文书)
| Template | Description | 语言 |
|----------|-------------|------|
| contract | 合同模板 | 中文 |
| authorization | 授权委托书模板 | 中文 |

### Academic Documents (学术论文)
| Template | Description | 语言 |
|----------|-------------|------|
| paper | 学术论文格式 | 中文 |
| thesis | 毕业论文模板 | 中文 |

### Custom (自定义模板)
| Template | Description | 语言 |
|----------|-------------|------|
| custom | 用户自定义模板 | 中文 |


## How to Create Templates

Create a Markdown file in `templates/` directory:

```markdown
# Document Title

[variables]
title: Default Title
author: Author Name
date: 2026-01-01

[content]
## Section 1
Your content here.

## Section 2
More content...
```

### Template Format

| Section | Purpose |
|---------|---------|
| `# Title` | Document title (appears as heading) |
| `[variables]` | Default variable values |
| `[content]` | Document body content |


## Usage Examples

### List All Templates

```bash
python document_generator.py -l
```

Output:
```
Available templates:
  - notice
  - notification
  - contract
  - paper
```

### Generate with Defaults

```bash
python document_generator.py notice -o output.docx
```

### Generate with Custom Variables

```bash
python document_generator.py notice \
  -o report.docx \
  -v title="Annual Report" \
  -v author="Finance Dept" \
  -v content="Quarterly financial summary..."
```


## Command Options

| Option | Description |
|--------|-------------|
| template | Template name (without .md) |
| -o, --output | Output filename (default: output.docx) |
| -l, --list | List available templates |
| -v, --variable | Add variable (key=value) |


## Add Custom Template

We provide a **custom template** for users to create their own documents:

| Template | Description |
|----------|-------------|
| custom | User-defined template (editable) |

### How to Use Custom Template

```bash
# Generate with default values
python document_generator.py custom -o mydoc.docx

# Generate with your own variables
python document_generator.py custom -o mydoc.docx \
  -v title="My Title" \
  -v author="My Name" \
  -v date="2026-02-10"
```

### Create Your Own Template

1. Copy `templates/custom/custom.md` to a new file
2. Edit the template content
3. Save with a new name (e.g., `templates/custom/my_template.md`)
4. Use your template:

```bash
python document_generator.py my_template -o output.docx
```

---

## Add New Template

1. Create a `.md` file in `templates/` subdirectory
2. Follow the template format
3. Test with:

```bash
python document_generator.py your_template -o test.docx
```


## Requirements

| Package | Version |
|---------|---------|
| python-docx | >=1.1.0 |


## License

MIT License - Free to use and modify


## Author

Created with Claude Code

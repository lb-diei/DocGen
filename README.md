# DocGen - Document Generator & Formatter

**Professional Document Automation Tool - Create and Format Perfect DOCX Files**

> *Your document automation assistant - Create and format perfect documents in seconds!*

**Author:** Your Name  
**License:** MIT

DocGen helps you:
- Generate professional documents from Word templates
- Format existing documents to standard styles (GB/T 9704-2012)
- Use GUI interface for easy document creation

---

## Core Features

### 1. Document Generator (Templates)
Generate documents from pre-designed Word templates with variables.

### 2. Document Formatter (Formatting)
Format any Word document to standard styles:
- Auto-detect heading levels
- Apply consistent fonts and sizes
- Set margins and line spacing

### 3. GUI Frontend (Graphical Interface)
Easy-to-use interface for:
- Selecting templates
- Customizing fonts, sizes, alignment
- Previewing and applying styles

---

## Quick Start

### Option 1: GUI Interface (Recommended)
```bash
python doc_gen_gui.py
```

### Option 2: Command Line - Document Generation
```bash
# Generate from template
python document_generator.py notice -o my_notice.docx

# With custom variables
python document_generator.py notice -o year_end.docx \
  -v title="2025 Annual Summary Notice" \
  -v author="HR Department"
```

### Option 3: Command Line - Document Formatting
```bash
# Format a Word document to standard style
python doc_formatter.py input.docx -o output.docx

# With markdown input
python doc_formatter.py content.md -o output.docx
```

---

## Project Structure

```
DocGen/
├── document_generator.py     # Main CLI - generate DOCX from templates
├── doc_formatter.py          # Format documents to standard styles
├── doc_gen_gui.py            # Graphical interface (recommended)
├── templates/                # 11 Word templates ready to use
│   ├── government/           # Government documents
│   │   ├── notice.docx       # Official notice template
│   │   └── request.docx     # Request document template
│   ├── enterprise/           # Business documents
│   │   ├── notification.docx
│   │   ├── meeting.docx
│   │   ├── report.docx
│   │   └── invitation.docx
│   ├── legal/               # Legal documents
│   │   ├── contract.docx
│   │   └── authorization.docx
│   ├── academic/            # Academic documents
│   │   ├── paper.docx
│   │   └── thesis.docx
│   └── custom/              # Your custom template
│       └── custom.docx      # Blank template - edit freely!
├── __init__.py
├── setup.py
└── README.md
```

---

## Available Templates (11 Templates)

### Government Documents
| Template | Description |
|----------|-------------|
| `notice` | Official notice for internal/external communication |
| `request` | Formal request document for approval |

### Enterprise Documents
| Template | Description |
|----------|-------------|
| `notification` | Internal company announcements |
| `meeting` | Meeting minutes with action items |
| `report` | Work reports with data tables |
| `invitation` | Event invitations with schedules |

### Legal Documents
| Template | Description |
|----------|-------------|
| `contract` | Standard contract with dual signatures |
| `authorization` | Authorization letters with agent details |

### Academic Documents
| Template | Description |
|----------|-------------|
| `paper` | Academic paper with abstract & references |
| `thesis` | Graduate thesis with all required sections |

### Custom Template
| Template | Description |
|----------|-------------|
| `custom` | Blank template for your own designs! |

---

## Installation

```bash
# Install dependencies
pip install python-docx

# Optional: Install as package
pip install -e .
```

---

## Creating Custom Templates

Create a Word document (.docx) in `templates/` directory with placeholders:

```
{{title}}     - Document title
{{author}}    - Author name
{{date}}      - Date
{{content}}   - Main content
{{variable}}  - Any custom variable
```

---

## License

MIT License - Free for personal and commercial use.

**Author:** Your Name  
**Repository:** https://github.com/lb-diei/DocGen

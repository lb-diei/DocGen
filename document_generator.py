#!/usr/bin/env python3
"""
DocGen - Document Generator

Generate professional DOCX documents from Word templates.
"""

import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Optional
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


class DocumentGenerator:
    """Generate DOCX documents from Word templates."""
    
    def __init__(self, template_dir: str = None):
        """Initialize with template directory."""
        if template_dir is None:
            # Default to templates folder relative to this script
            script_dir = Path(__file__).parent
            template_dir = str(script_dir / "templates")
        self.template_dir = Path(template_dir)
        self.default_vars = {
            "title": "文章标题",
            "author": "作者姓名",
            "date": datetime.now().strftime("%Y-%m-%d"),
            "introduction": "引言/开篇内容...",
            "heading1": "一级标题",
            "heading2": "二级标题",
            "heading3": "三级标题",
            "content": "正文内容...",
            "conclusion": "结论/总结...",
            "signature": "落款（签名、日期）",
        }
    
    def load_template(self, template_name: str) -> Document:
        """Load a Word template file."""
        # Search in all subdirectories
        for template_path in self.template_dir.rglob(f"{template_name}.docx"):
            if template_path.is_file():
                return Document(template_path)
        
        raise FileNotFoundError(f"Template not found: {template_name}.docx")
    
    def replace_variables(self, doc: Document, variables: Dict[str, Any]) -> Document:
        """Replace variables in the document."""
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                text = run.text
                for key, value in variables.items():
                    # Replace {{key}} format
                    placeholder = "{{" + key + "}}"
                    if placeholder in text:
                        text = text.replace(placeholder, str(value))
                    # Replace {{ key }} format with spaces
                    placeholder_spaced = "{{ " + key + " }}"
                    if placeholder_spaced in text:
                        text = text.replace(placeholder_spaced, str(value))
                run.text = text
        return doc
    
    def generate_document(
        self,
        template_name: str,
        output_name: str,
        variables: Optional[Dict[str, Any]] = None,
    ) -> str:
        """Generate a DOCX document from a template."""
        # Load template
        doc = self.load_template(template_name)
        
        # Prepare variables
        template_vars = self.default_vars.copy()
        if variables:
            template_vars.update(variables)
        
        # Replace variables
        doc = self.replace_variables(doc, template_vars)
        
        # Save document
        output_path = Path(output_name)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output_path))
        
        return str(output_path)
    
    def list_templates(self) -> list:
        """List available templates."""
        if not self.template_dir.exists():
            return []
        
        templates = []
        # Search recursively for all .docx files
        for f in self.template_dir.rglob('*.docx'):
            # Get relative path and use that as template name
            rel_path = f.relative_to(self.template_dir)
            template_name = str(rel_path).replace('\\', '/').replace('/', '_')
            templates.append(template_name)
        
        return sorted(set(templates))


def main():
    """CLI interface."""
    import argparse
    
    parser = argparse.ArgumentParser(description='DocGen - Generate DOCX from templates')
    parser.add_argument('template', nargs='?', help='Template name (without .docx extension)')
    parser.add_argument('-o', '--output', default='output.docx', help='Output filename')
    parser.add_argument('-l', '--list', action='store_true', help='List available templates')
    parser.add_argument('-v', '--variable', action='append', help='Variable in format key=value')
    
    args = parser.parse_args()
    
    generator = DocumentGenerator()
    
    if args.list:
        templates = generator.list_templates()
        if templates:
            print("Available templates:")
            for t in templates:
                print(f"  - {t}")
        else:
            print("No templates found in 'templates/' directory.")
        return
    
    if not args.template:
        parser.print_help()
        return
    
    # Parse variables from command line
    variables = {}
    if args.variable:
        for var in args.variable:
            if '=' in var:
                key, value = var.split('=', 1)
                variables[key] = value
    else:
        # Interactive mode: prompt user for variables
        print(f"\n[DocGen] Generating document from template: {args.template}")
        print("Please fill in the following values (press Enter to use default):\n")
        # Writing variables
        common_vars = [
            'title',        # 文章标题
            'author',       # 作者
            'date',         # 日期
            'introduction', # 引言
            'heading1',     # 一级标题
            'heading2',     # 二级标题
            'heading3',     # 三级标题
            'content',      # 正文内容
            'conclusion',   # 结论
            'signature',    # 落款
        ]
        for var in common_vars:
            default = generator.default_vars.get(var, '')
            user_input = input(f"  {var} [{default}]: ").strip()
            if user_input:
                variables[var] = user_input
            elif default:
                variables[var] = default
        print()
    
    # Generate document
    try:
        output_path = generator.generate_document(
            args.template,
            args.output,
            variables,
        )
        print(f"Document generated: {output_path}")
    except Exception as e:
        print(f"Error: {e}")


if __name__ == '__main__':
    main()

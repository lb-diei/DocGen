#!/usr/bin/env python3
"""
DocFormatter - Document Template System

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
    
    def __init__(self, template_dir: str = "templates"):
        """Initialize with template directory."""
        self.template_dir = Path(template_dir)
        self.default_vars = {
            "title": "Document Title",
            "author": "Author Name",
            "date": datetime.now().strftime("%Y-%m-%d"),
            "content": "Your content here...",
        }
    
    def load_template(self, template_name: str) -> Document:
        """Load a Word template file."""
        template_path = self.template_dir / f"{template_name}.docx"
        
        if not template_path.exists():
            raise FileNotFoundError(f"Template not found: {template_path}")
        
        return Document(template_path)
    
    def replace_variables(self, doc: Document, variables: Dict[str, Any]) -> Document:
        """Replace variables in the document."""
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                text = run.text
                for key, value in variables.items():
                    if f"{{{{ {key} }}}}" in text:
                        text = text.replace(f"{{{{ {key} }}}}", str(value))
                        text = text.replace(f"{{{{{key}}}}", str(value))
                        text = text.replace(f"{{ {key} }}", str(value))
                    elif f"{{{{{key}}}}}" in text:
                        text = text.replace(f"{{{{{key}}}}}", str(value))
                    elif f"{{{{{key}}}}}" in text:
                        text = text.replace(f"{{{{{key}}}}}", str(value))
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
        for f in self.template_dir.glob('*.docx'):
            templates.append(f.stem)
        
        return templates


def main():
    """CLI interface."""
    import argparse
    
    parser = argparse.ArgumentParser(description='DocFormatter - Generate DOCX from templates')
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
    
    # Parse variables
    variables = {}
    if args.variable:
        for var in args.variable:
            if '=' in var:
                key, value = var.split('=', 1)
                variables[key] = value
    
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

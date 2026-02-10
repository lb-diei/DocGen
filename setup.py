"""
DocGen Setup Configuration
"""

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="docgen",
    version="1.0.0",
    author="Your Name",
    description="Document Generator & Formatter - Generate and format professional DOCX files",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/lb-diei/DocGen",
    packages=find_packages(),
    python_requires=">=3.7",
    install_requires=[
        "python-docx>=0.8.11",
    ],
    extras_require={
        "gui": [
            "tkinter",
        ],
    },
    entry_points={
        "console_scripts": [
            "docgen=docgen.document_generator:main",
            "docfmt=docgen.doc_formatter:main",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
    ],
)

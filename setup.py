#!/usr/bin/env python3
"""
Setup script for Excel-to-LLM Context Feeder Tool.
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read the README file
this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text()

setup(
    name="excel-llm-tool",
    version="1.0.0",
    description="Convert Excel files to LLM-friendly text formats with Streamlit GUI",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="Excel-to-LLM Tool Team",
    author_email="",
    url="",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Text Processing :: Markup",
        "Topic :: Office/Business",
    ],
    python_requires=">=3.8",
    install_requires=[
        "pandas>=2.2.2",
        "openpyxl>=3.1.5",
        "streamlit>=1.38.0",
        "numpy>=1.21.0",
    ],
    extras_require={
        "dev": [
            "pytest>=6.0",
            "pytest-cov>=2.0",
            "black>=21.0",
            "flake8>=3.8",
            "mypy>=0.800",
        ],
    },
    entry_points={
        "console_scripts": [
            "excel-llm-tool=src.cli:main",
        ],
    },
    include_package_data=True,
    zip_safe=False,
    keywords="excel, llm, csv, json, markdown, streamlit, data-conversion",
    project_urls={
        "Bug Reports": "",
        "Source": "",
        "Documentation": "",
    },
)

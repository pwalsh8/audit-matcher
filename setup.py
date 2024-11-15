"""
Setup script for audit-matcher project
Creates project structure and installs dependencies
"""
import os
import subprocess
import sys
from pathlib import Path

def create_project_structure():
    """Create the project directories and files"""
    # Create main directories
    directories = [
        'src',
        'tests',
        'data',
    ]

    for dir_name in directories:
        Path(dir_name).mkdir(parents=True, exist_ok=True)

    # Create src files
    src_files = {
        'src/__init__.py': '',
        'src/main.py': '''import streamlit as st

def main():
    st.title("Audit Matcher")
    st.write("Welcome to Audit Matcher!")

    # File uploaders
    selections_file = st.file_uploader("Upload Selections Excel", type=['xlsx'])
    pdf_files = st.file_uploader("Upload PDFs", accept_multiple_files=True, type=['pdf'])

if __name__ == "__main__":
    main()
''',
        'src/matcher.py': '''import pdfplumber
import pandas as pd

def process_pdf(file_path):
    """Extract text from PDF file"""
    with pdfplumber.open(file_path) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text() + "\\n"
    return text
''',
        'src/utils.py': '''from typing import Union
from decimal import Decimal, InvalidOperation
from pathlib import Path
import pandas as pd
import streamlit as st
import tempfile

class CommonUtils:
    """Shared utility functions"""
    # Add shared utility methods
'''
    }

    for file_path, content in src_files.items():
        Path(file_path).parent.mkdir(parents=True, exist_ok=True)
        with open(file_path, 'w') as f:
            f.write(content)

    # Create requirements.txt
    requirements = [
        'streamlit',
        'pandas',
        'pdfplumber',
        'openpyxl',
        'python-dotenv',
        'PyPDF2'  # Add PyPDF2 to the requirements
    ]

    with open('requirements.txt', 'w') as f:
        f.write('\n'.join(requirements))

def setup_dependencies():
    """Install system dependencies"""
    # Install poppler for PDF processing
    if os.name == 'nt':  # Windows
        print("Please install poppler manually from: https://github.com/oschwartz10612/poppler-windows/releases")
        print("Then add the bin directory to your system PATH")
    else:  # Unix/Mac
        try:
            subprocess.run(['apt-get', 'install', '-y', 'poppler-utils'], check=True)
        except Exception:
            print("Please install poppler-utils manually using your system's package manager")

def create_virtual_environment():
    """Create and activate virtual environment"""
    if not Path('venv').exists():
        subprocess.run([sys.executable, '-m', 'venv', 'venv'], check=True)

    # Install requirements
    if os.name == 'nt':  # Windows
        pip_path = 'venv\\Scripts\\pip'
    else:  # Unix/Mac
        pip_path = 'venv/bin/pip'

    subprocess.run([pip_path, 'install', '-r', 'requirements.txt'], check=True)

def main():
    """Main setup function"""
    print("Setting up Audit Matcher project...")
    create_project_structure()
    setup_dependencies()
    create_virtual_environment()
    print("\nSetup complete! To start the application:")
    print("1. Activate virtual environment:")
    if os.name == 'nt':  # Windows
        print("   .\\venv\\Scripts\\activate")
    else:  # Unix/Mac
        print("   source venv/bin/activate")
    print("2. Run the application:")
    print("   streamlit run src/main.py")

if __name__ == "__main__":
    main()

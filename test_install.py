import sys
from pathlib import Path

def check_dependencies():
    """Check if all required dependencies are installed and working"""
    failures = []
    
    # Test Python packages
    packages = {
        'pdf2image': 'Convert PDFs to images',
        'pdfplumber': 'Extract text from PDFs',
        'openpyxl': 'Excel file handling',
        'streamlit': 'Web interface',
        'pandas': 'Data processing',
        'PIL': 'Image processing',
        'thefuzz': 'String matching',
        'Levenshtein': 'String distance calculations'
    }
    
    print("Checking Python packages...")
    for package, description in packages.items():
        try:
            if package == 'PIL':
                __import__('PIL')
            else:
                __import__(package)
            print(f"✓ {package:<15} - {description}")
        except ImportError as e:
            failures.append(f"✗ {package} - {str(e)}")

    # Check Poppler installation
    print("\nChecking Poppler installation...")
    poppler_path = Path(r'C:\Program Files\poppler\Library\bin')
    if (poppler_path.exists()):
        print(f"✓ Poppler found at: {poppler_path}")
    else:
        failures.append("✗ Poppler not found. Please install from: https://github.com/oschwartz10612/poppler-windows/releases")

    # Report results
    if failures:
        print("\nMissing dependencies:")
        for failure in failures:
            print(failure)
        print("\nPlease install missing dependencies using:")
        print("pip install pdf2image pdfplumber openpyxl streamlit pandas pillow thefuzz[speedup] python-Levenshtein")
        print("\nAnd install Poppler from: https://github.com/oschwartz10612/poppler-windows/releases")
        return False
    else:
        print("\nAll dependencies installed successfully!")
        return True

if __name__ == "__main__":
    check_dependencies()
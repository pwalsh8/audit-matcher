
"""Audit Matcher Tool - Matches audit selections with PDF documents"""

import streamlit as st
import pandas as pd
import pdfplumber
import re
from typing import List, Dict, Any, Union, Optional
import logging
from decimal import Decimal, InvalidOperation
from pathlib import Path
import os
import tempfile
from pdf2image import convert_from_path
from PIL import Image
import openpyxl
from openpyxl.drawing.image import Image as OpenPyxlImage
import shutil
import time

# Constants and Configuration
POPPLER_PATH = r'C:\Program Files\poppler\poppler-24.08.0\Library\bin'
OUTPUT_DIR = Path("output_images")

# Configure logging
logging.basicConfig(
    filename='app.log',
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# Utility Classes & Functions 
class PDFHandler:
    """PDF processing functionality"""
    def __init__(self):
        self.output_dir = ensure_output_directory()
        self.temp_dir = self.output_dir / "temp"

    def extract_text(self, pdf_file: Union[str, Path, st.runtime.uploaded_file_manager.UploadedFile]) -> str:
        """Extract text from PDF file"""
        try:
            file_path = CommonUtils.get_file_path(pdf_file)
            with pdfplumber.open(file_path) as pdf:
                text = ""
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
                return text
        except Exception as e:
            logger.error(f"Error extracting text from PDF: {str(e)}")
            return ""
        finally:
            if isinstance(pdf_file, st.runtime.uploaded_file_manager.UploadedFile):
                file_path.unlink(missing_ok=True)

    def convert_to_images(self, pdf_file: Union[str, Path, st.runtime.uploaded_file_manager.UploadedFile], 
                         selection_id: str) -> List[str]:
        """Convert PDF pages to images"""
        try:
            output_dir = ensure_output_directory()
            file_path = CommonUtils.get_file_path(pdf_file)
            
            # Create safe selection ID for filenames
            safe_id = str(selection_id).replace('.', '_').replace(' ', '_')
            
            # Convert PDF to images
            images = convert_from_path(file_path, poppler_path=POPPLER_PATH)
            
            image_paths = []
            for i, image in enumerate(images):
                image_path = output_dir / f"{safe_id}_page_{i+1}.png"
                image.save(str(image_path.resolve()), "PNG")
                logger.info(f"Saved PDF page {i+1} to: {image_path}")
                image_paths.append(str(image_path.resolve()))
            
            return image_paths
        except Exception as e:
            logger.error(f"Error converting PDF to images: {str(e)}")
            return []
        finally:
            if isinstance(pdf_file, st.runtime.uploaded_file_manager.UploadedFile):
                try:
                    file_path.unlink(missing_ok=True)
                except Exception as e:
                    logger.error(f"Error cleaning up temp file: {e}")

    def cleanup(self) -> None:
        """Clean up temporary files and images"""
        try:
            if self.temp_dir.exists():
                shutil.rmtree(self.temp_dir)
            self.temp_dir.mkdir(parents=True, exist_ok=True)
            
            # Only remove PNG files from output_dir
            for f in self.output_dir.glob("*.png"):
                f.unlink()
        except Exception as e:
            logger.error(f"Error during cleanup: {e}")

class TextMatcher:
    """Advanced text matching utility with flexible string matching"""
    def __init__(self):
        self._word_separators = r'[\s\-_,.]+'

    def ratio(self, str1: str, str2: str) -> float:
        """Get similarity ratio between two strings"""
        if not str1 or not str2:
            return 0.0
            
        str1, str2 = str1.lower(), str2.lower()
        
        # Exact match
        if str1 == str2:
            return 100.0
            
        # Prepare word sets
        words1 = set(re.split(self._word_separators, str1))
        words2 = set(re.split(self._word_separators, str2))
        
        # Remove empty strings
        words1.discard('')
        words2.discard('')
        
        if not words1 or not words2:
            return 0.0
            
        # Calculate word overlap
        common_words = words1 & words2
        total_words = max(len(words1), len(words2))
        
        if not total_words:
            return 0.0
            
        return (len(common_words) / total_words) * 100.0

# Initialize text matcher
fuzz = TextMatcher()

def ensure_output_directory() -> Path:
    """Ensure output directory exists and return absolute Path"""
    # Get absolute base path - this should be project root
    base_dir = Path.cwd()
    if 'src' in base_dir.parts:
        base_dir = base_dir.parent
    
    output_dir = base_dir / "output_images"
    temp_dir = output_dir / "temp"
    
    # Create both directories if they don't exist
    output_dir.mkdir(parents=True, exist_ok=True)
    temp_dir.mkdir(parents=True, exist_ok=True)
    
    # Return resolved absolute path
    return output_dir.resolve()

class CommonUtils:
    """Shared utility functions used across modules"""
    @staticmethod
    def format_currency(amount: Union[float, str, Decimal]) -> str:
        try:
            if isinstance(amount, str):
                amount = float(amount.replace(',', ''))
            return f"${Decimal(str(amount)):,.2f}"
        except (InvalidOperation, ValueError) as e:
            log_error(f"Error formatting currency: {str(e)}")
            return "Invalid Amount"

    @staticmethod
    def normalize_amount(amount: Union[str, float, Decimal]) -> Decimal:
        try:
            if pd.isna(amount):
                return Decimal('0')
            if isinstance(amount, str):
                cleaned = ''.join(c for c in amount if c.isdigit() or c in '.-')
                amount = float(cleaned)
            return Decimal(str(float(amount))).quantize(Decimal('0.01'))
        except (InvalidOperation, ValueError) as e:
            log_error(f"Error normalizing amount {amount}: {str(e)}")
            return Decimal('0')

    @staticmethod
    def get_file_path(file: Union[str, Path, st.runtime.uploaded_file_manager.UploadedFile]) -> Path:
        if isinstance(file, st.runtime.uploaded_file_manager.UploadedFile):
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                tmp.write(file.getvalue())
                return Path(tmp.name)
        return Path(file)

    @staticmethod
    def log_error(message: str, exception: Optional[Exception] = None) -> None:
        """Centralized error logging"""
        logger.error(message)
        if exception:
            logger.error(f"Exception details: {str(exception)}", exc_info=True)

# Core Processing Functions
def extract_primary_amount(text: str) -> Optional[float]:
    """
    Extract and prioritize invoice total amounts from text.
    """
    # Enhanced patterns for invoice totals with stronger contextual matching
    total_patterns = [
        r'(?i)(?:total\s+due|invoice\s+total|grand\s+total|total\s+amount)[\s:]*[\$]?\s*([\d,]+\.?\d*)',
        r'(?i)balance\s+due[\s:]*[\$]?\s*([\d,]+\.?\d*)',
        r'(?i)amount\s+due[\s:]*[\$]?\s*([\d,]+\.?\d*)',
        r'(?i)total[\s:]*[\$]?\s*([\d,]+\.?\d*)(?!\s*\w)',  # Total not followed by other words
        r'\$\s*([\d,]+\.?\d*)(?=\s*(?:total|due|usd))',  # Amount followed by total/due/usd
        r'\$\s*([\d,]+\.?\d*)(?=\s*$)',  # Dollar amount at end of line
    ]
    
    # Normalize text: remove extra spaces and line breaks
    text = ' '.join(text.split())
    logger.debug(f"Processing text for primary amount: {text[:200]}...")
    
    # First pass: Look for amounts with strongest contextual indicators
    for pattern in total_patterns:
        matches = re.finditer(pattern, text)
        amounts = []
        for match in matches:
            try:
                amount_str = match.group(1).replace(',', '')
                amount = float(amount_str)
                if amount > 0:
                    # Verify this isn't just an invoice number
                    pre_context = text[max(0, match.start() - 50):match.start()]
                    if not re.search(r'(?i)invoice\s+(?:no|number|#)', pre_context):
                        amounts.append(amount)
            except (ValueError, IndexError):
                continue
        
        if amounts:
            # If multiple amounts found, prefer the largest that isn't an obvious invoice number
            valid_amounts = [amt for amt in amounts if amt > 1000]  # Assume amounts under 1000 might be invoice numbers
            if valid_amounts:
                primary_amount = max(valid_amounts)
                logger.debug(f"Found primary amount {primary_amount} using pattern: {pattern}")
                return primary_amount
    
    # Second pass: Extract all potential amounts and apply heuristics
    all_amounts = extract_amounts_from_text(text)
    if all_amounts:
        # Filter out likely invoice numbers and small amounts
        valid_amounts = [amt for amt in all_amounts if amt > 1000]
        if valid_amounts:
            # Take the largest valid amount
            primary_amount = max(valid_amounts)
            logger.debug(f"Using fallback: largest valid amount found {primary_amount}")
            return primary_amount
    
    return None

def extract_amounts_from_text(text: str) -> List[float]:
    """
    Extract all potential currency amounts from text.
    """
    # Enhanced patterns for amount extraction with better context
    patterns = [
        r'(?i)(?:total|amount|balance|due).*?[\$]?\s*([\d,]+\.?\d*)',
import pdfplumber
import re
from typing import List, Dict, Any, Union, Optional
import logging
from decimal import Decimal, InvalidOperation
from pathlib import Path
import pandas as pd
import streamlit as st
from pdf2image import convert_from_path
import os
import tempfile
from utils import CommonUtils, ensure_output_directory, PDFHandler

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

logger = logging.getLogger(__name__)

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
        r'(?i)[\$]?\s*([\d,]+\.?\d*)(?=\s*(?:usd|dollars|total|due))',
        r'\$\s*([\d,]+\.?\d*)',  # Standard dollar amounts
        r'(?<!\w)([\d,]+\.\d{2})(?!\w)'  # Standalone decimal numbers with exactly 2 decimal places
    ]
    
    amounts = []
    seen = set()
    
    for pattern in patterns:
        matches = re.finditer(pattern, text)
        for match in matches:
            try:
                amount_str = match.group(1).replace(',', '')
                amount = float(amount_str)
                
                # Additional validation
                if amount > 0 and amount not in seen:
                    # Check context to avoid invoice numbers
                    pre_context = text[max(0, match.start() - 50):match.start()]
                    if not re.search(r'(?i)invoice\s+(?:no|number|#)', pre_context):
                        amounts.append(amount)
                        seen.add(amount)
            except (ValueError, IndexError):
                continue
    
    logger.debug(f"Extracted amounts from text: {amounts}")
    return amounts

def extract_text_from_pdf(pdf_file: Union[str, Path, st.runtime.uploaded_file_manager.UploadedFile]) -> str:
    pdf_handler = PDFHandler()
    return pdf_handler.extract_text(pdf_file)

def convert_pdf_to_images(pdf_file: Union[str, Path, st.runtime.uploaded_file_manager.UploadedFile], selection_id: str) -> List[str]:
    """Convert PDF pages to images"""
    pdf_handler = PDFHandler()
    output_dir = ensure_output_directory()  # Get output directory path
    
    try:
        file_path = pdf_handler.get_file_path(pdf_file)
        images = convert_from_path(file_path, poppler_path=r'C:\Program Files\poppler\poppler-24.08.0\Library\bin')

        image_paths = []
        for i, image in enumerate(images):
            # Ensure valid filename by removing any decimal points from selection_id
            safe_id = str(selection_id).replace('.', '_')
            image_path = output_dir / f"{safe_id}_page_{i+1}.png"
            image.save(str(image_path), "PNG")
            logger.info(f"Saved PDF page {i+1} to: {image_path}")
            
            # Verify the image was created
            if not image_path.exists():
                raise FileNotFoundError(f"Failed to save image to {image_path}")
                
            image_paths.append(str(image_path))
            
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

def calculate_match_score(pdf_amount: Union[str, float, Decimal], selection_amount: Union[str, float, Decimal]) -> float:
    """Calculate a match score between two amounts."""
    try:
        # Normalize both amounts to Decimal with 2 decimal places
        pdf_decimal = normalize_amount(pdf_amount)
        selection_decimal = normalize_amount(selection_amount)
        
        logger.debug(f"Comparing amounts - PDF: {pdf_decimal} vs Selection: {selection_decimal}")
        
        if pdf_decimal == selection_decimal:
            logger.debug("Exact match found!")
            return 1.0
        
        # Calculate percentage difference
        if selection_decimal == Decimal('0'):
            return 0.0
            
        diff = abs(pdf_decimal - selection_decimal) / selection_decimal
        if diff <= 0.01:  # Within 1%
            return 0.9
        elif diff <= 0.05:  # Within 5%
            return 0.7
        elif diff <= 0.1:  # Within 10%
            return 0.5
        return 0.0
    except Exception as e:
        logger.error(f"Error calculating match score: {str(e)}")
        return 0.0

def match_entries(selections: List[Dict], pdf_entries: List[Dict], threshold: float = 0.8) -> List[Dict]:
    """Match entries between selections and PDF data using multiple criteria."""
    matches = []
    
    # Log amounts for debugging
    for selection in selections:
        selection_amount = normalize_amount(selection[selection['amount_column']])
        logger.debug(f"Selection amount (normalized): {selection_amount}")
        
        potential_matches = []
        for pdf_entry in pdf_entries:
            pdf_amount = normalize_amount(pdf_entry['amount'])
            logger.debug(f"PDF amount (normalized): {pdf_amount} from {pdf_entry['pdf_name']}")
            
            # Compare normalized amounts
            if pdf_amount == selection_amount or abs(pdf_amount - selection_amount) <= Decimal('0.01'):
                match_score = 100  # Base score for amount match
                similarity = 0  # Initialize similarity score
                
                # Extract customer and date info from PDF text if available
                pdf_text = pdf_entry.get('text', '').lower()
                selection_customer = str(selection.get('customer', '')).lower()
                
                # Customer name fuzzy matching if available
                if selection_customer and pdf_text:
                    similarity = fuzz.ratio(selection_customer, pdf_text)
                    if similarity > threshold * 100:
                        match_score += 20
                
                # Date matching if available
                if 'date' in selection and pdf_entry.get('date') and selection['date'] == pdf_entry['date']:
                        match_score += 10
                
                potential_matches.append({
                    'pdf_entry': pdf_entry,
                    'score': match_score,
                    'similarity': similarity
                })
        
        # Sort potential matches by score
        potential_matches.sort(key=lambda x: (x['score'], x.get('similarity', 0)), reverse=True)
        
        if potential_matches:
            best_match = potential_matches[0]
            matches.append({
                'selection': selection,
                'match': best_match['pdf_entry'],
                'score': best_match['score'],
                'similarity': best_match.get('similarity', 0)
            })
        else:
            matches.append({
                'selection': selection,
                'match': None,
                'score': 0,
                'similarity': 0
            })
    
    return matches

def match_documents(df: pd.DataFrame, unique_id_column: str, amount_column: str, 
                   pdf_files: List[Union[str, Path, st.runtime.uploaded_file_manager.UploadedFile]], 
                   threshold: float = 0.5) -> List[Dict[str, Any]]:
    """Match amounts from Excel selections with amounts found in PDFs."""
    logger.debug(f"Processing {len(df)} selections and {len(pdf_files)} PDFs")
    
    # Prepare selections data
    selections = []
    for _, row in df.iterrows():
        try:
            # Clean and normalize amount
            amount_str = str(row[amount_column]).replace('$', '').replace(',', '').strip()
            amount = float(amount_str)
            
            selections.append({
                'id': str(row[unique_id_column]),
                'amount': amount,
                'raw_data': row.to_dict()
            })
            logger.debug(f"Processed selection - ID: {row[unique_id_column]}, Amount: {amount}")
        except Exception as e:
            logger.error(f"Error processing selection {row[unique_id_column]}: {e}")
            continue

    # Process PDFs and extract amounts
    pdf_entries = []
    output_dir = ensure_output_directory()
    
    for pdf_file in pdf_files:
        try:
            text = extract_text_from_pdf(pdf_file)
            amounts = extract_amounts_from_text(text)
            logger.debug(f"Found amounts in {pdf_file.name}: {amounts}")
            
            # Store each amount found in the PDF
            for amount in amounts:
                pdf_entries.append({
                    'file': pdf_file,
                    'amount': amount,
                    'text': text,
                    'pdf_name': pdf_file.name,
                })
        except Exception as e:
            logger.error(f"Error processing PDF {pdf_file.name}: {e}")
            continue

    # Match selections with PDFs
    all_matches = []
    
    # Use the output directory consistently
    output_dir = ensure_output_directory()
    
    for selection in selections:
        selection_matches = []
        
        for pdf_entry in pdf_entries:
            try:
                # Compare amounts with small tolerance
                amount_diff = abs(selection['amount'] - pdf_entry['amount'])
                if amount_diff < 0.01:  # 1 cent tolerance
                    # Generate safe ID for filenames
                    safe_id = str(selection['id']).replace('.', '_').replace(' ', '_')
                    
                    # Convert PDF to images and ensure paths are absolute
                    image_paths = convert_pdf_to_images(pdf_entry['file'], safe_id)
                    image_paths = [str(Path(p).resolve()) for p in image_paths]
                    
                    match_info = {
                        'Selection ID': safe_id,
                        'Selection Data': selection['raw_data'],
                        'Selection Amount': format_currency(selection['amount']),
                        'PDF Name': pdf_entry['pdf_name'],
                        'PDF Amount': format_currency(pdf_entry['amount']),
                        'PDF Text': pdf_entry['text'],
                        'Match Type': 'Exact',
                        'Match Score': 100 - (amount_diff * 100),
                        'Matched Pages': image_paths
                    }
                    selection_matches.append(match_info)
                    logger.info(f"Found match for selection {safe_id}")
            except Exception as e:
                logger.error(f"Error matching selection {selection['id']}: {e}")
                continue
        
        # Add all matches for this selection
        if selection_matches:
            all_matches.extend(selection_matches)
        else:
            logger.warning(f"No matches found for selection {selection['id']}")
    
    logger.info(f"Found {len(all_matches)} total matches")
    return all_matches

def verify_match_accuracy(match: Dict[str, Any]) -> float:
    """
    Verify the accuracy of a match (placeholder for future enhancement).
    
    Args:
        match: Dictionary containing match details
        
    Returns:
        float: Confidence score between 0 and 1
    """
    # This could be enhanced with more sophisticated matching logic
    # For now, return 1.0 for exact matches
    return 1.0 if match else 0.0

def extract_amount_from_pdf(pdf_file: Union[str, Path, st.runtime.uploaded_file_manager.UploadedFile]) -> Optional[float]:
    """
    Extract the most likely amount from a PDF file.
    """
    try:
        text = extract_text_from_pdf(pdf_file)
        primary_amount = extract_primary_amount(text)
        return primary_amount
    except Exception as e:
        logger.error(f"Error extracting amount from PDF: {str(e)}")
        return None

def interactive_matching(df: pd.DataFrame, pdf_files: List, unique_id_column: str, amount_column: str):
    """Interactive matching process with preview and confirmation"""
    try:
        # Initialize matching state if not exists
        if 'matching_state' not in st.session_state:
            st.session_state.matching_state = {
                'matches': [],
                'current_index': 0,
                'confirmed_matches': [],
                'skipped_matches': [],
                'completed': False
            }
            
            # Process initial matches
            result = process_files(df, unique_id_column, amount_column, pdf_files)
            if "error" in result:
                st.error(f"An error occurred: {result['error']}")
                return
                
            st.session_state.matching_state['matches'] = result.get("matches", [])
            logger.debug(f"Found {len(st.session_state.matching_state['matches'])} matches")

        state = st.session_state.matching_state
        
        # Display current match if available
        if not state['completed'] and len(state['matches']) > 0:
            current_match = state['matches'][state['current_index']]
            
            # Show match preview in columns
            col1, col2 = st.columns(2)
            with col1:
                st.write("### Selection Details")
                st.write(current_match['selection'])
            
            with col2:
                st.write("### PDF Details")
                st.write(current_match['match'])
            
            # Action buttons
            act1, act2, act3 = st.columns(3)
            with act1:
                if st.button("Skip", key=f"skip_{state['current_index']}"):
                    state['skipped_matches'].append(current_match)
                    _advance_match(state)
            
            with act2:
                if st.button("Confirm", key=f"confirm_{state['current_index']}"):
                    state['confirmed_matches'].append(current_match)
                    _advance_match(state)
            
            with act3:
                if st.button("Reject", key=f"reject_{state['current_index']}"):
                    _advance_match(state)
            
            # Show progress
            st.progress(state['current_index'] / len(state['matches']))
            st.write(f"Match {state['current_index'] + 1} of {len(state['matches'])}")
            
        elif state['completed']:
            st.write("Matching process completed.")
            st.write(f"Confirmed matches: {len(state['confirmed_matches'])}")
            st.write(f"Skipped matches: {len(state['skipped_matches'])}")

    except Exception as e:
        logger.error(f"Error in interactive matching: {str(e)}")
        st.error("An error occurred during matching. Please try again.")

def _advance_match(state: Dict):
    """Helper to advance to next match"""
    state['current_index'] += 1
    if state['current_index'] >= len(state['matches']):
        state['completed'] = True
    st.experimental_rerun()

def process_files(df: pd.DataFrame, unique_id_column: str, amount_column: str, pdf_files: List) -> Dict[str, Any]:
    """Process files and find matches"""
    try:
        logger.debug(f"Processing {len(pdf_files)} PDFs against {len(df)} selections")
        
        if df.empty:
            raise ValueError("DataFrame is empty")
            
        # Convert amounts to numeric and validate
        try:
            df[amount_column] = pd.to_numeric(df[amount_column], errors='coerce')
            if df[amount_column].isna().any():
                invalid_rows = df[df[amount_column].isna()].index.tolist()
                raise ValueError(f"Invalid amounts found in rows: {invalid_rows}")
        except Exception as e:
            raise ValueError(f"Error processing amounts: {str(e)}")

        # Find matches
        matches = match_documents(df, unique_id_column, amount_column, pdf_files)
        logger.info(f"Found {len(matches)} potential matches")
        
        return {
            "matches": matches,
            "total": len(matches),
            "status": "success"
        }
        
    except Exception as e:
        logger.error(f"Error processing files: {str(e)}")
        return {"error": str(e), "status": "error"}

# Replace direct function definitions with CommonUtils
normalize_amount = CommonUtils.normalize_amount
get_file_path = CommonUtils.get_file_path
format_currency = CommonUtils.format_currency
log_error = CommonUtils.log_error

__all__ = [
    'extract_amounts_from_text',
    'extract_primary_amount',
    'extract_text_from_pdf',
    'match_documents',
    'process_files'
]

# ...existing code...
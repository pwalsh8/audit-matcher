import streamlit as st
import pandas as pd
from matcher import (match_documents, extract_text_from_pdf, extract_primary_amount, 
                    process_files, extract_amounts_from_text)
from utils import (save_matches_to_excel, cleanup_output_images, load_and_validate_excel, 
                  PDFHandler, CommonUtils)
import logging
import tempfile
import time
from typing import List, Dict, Any
from pathlib import Path
from PIL import Image
import io  # Add missing import
from pdf2image import convert_from_path
from openpyxl import Workbook
import openpyxl
from openpyxl.drawing.image import Image as ExcelImage

# Configure logging with a more standardized setup
logging.basicConfig(
    filename='app.log',
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# Remove duplicate function and use CommonUtils directly
format_currency = CommonUtils.format_currency
get_file_path = CommonUtils.get_file_path
normalize_amount = CommonUtils.normalize_amount

def handle_error(error_message):
    """
    Handle errors and logging.
    
    Args:
        error_message: The error message to be logged and displayed
    """
    logger.error(error_message)
    st.error(error_message)
    st.write("Please check your files and try again")

def show_match_confirmation(matches):
    """
    Display interface for confirming matches when multiple possibilities exist.
    """
    if not matches:
        st.warning("No matches found in the provided files.")
        return []
        
    confirmed_matches = []
    
    # Group matches by Selection ID
    grouped_matches = {}
    for match in matches:
        selection_id = match['Selection ID']
        if selection_id not in grouped_matches:
            grouped_matches[selection_id] = []
        grouped_matches[selection_id].append(match)
    
    st.write("### Confirm Matches")
    st.write("Please review and confirm the matches for each selection.")
    
    for selection_id, potential_matches in grouped_matches.items():
        st.write(f"#### Selection {selection_id}")
        st.write(f"Amount: {potential_matches[0]['Selection Amount']}")
        
        if len(potential_matches) > 1:
            options = [f"{m['PDF Name']} - {m['PDF Amount']} ({m['Match Type']})" 
                      for m in potential_matches]
            selected = st.radio(
                f"Choose the correct match for Selection {selection_id}:",
                options,
                key=f"match_{selection_id}"
            )
            
            # Find the selected match
            selected_index = options.index(selected)
            confirmed_matches.append(potential_matches[selected_index])
        else:
            st.write(f"Single match found: {potential_matches[0]['PDF Name']}")
            confirmed_matches.append(potential_matches[0])
    
    # Add confirmation button
    if st.button("Confirm Matches"):
        if confirmed_matches:
            st.success("Matches confirmed successfully!")
            return confirmed_matches
        else:
            st.warning("Please select at least one match before confirming.")
            return []
            
    return None  # Return None if confirmation button not yet clicked

def save_uploaded_pdf(pdf_file) -> Path:
    """Save uploaded PDF to temporary file for processing"""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
            tmp.write(pdf_file.getvalue())
            return Path(tmp.name)
    except Exception as e:
        logger.error(f"Error saving PDF file: {str(e)}")
        raise

def preview_pdf_amounts(pdf_files: List) -> pd.DataFrame:
    """Extract and preview amounts from uploaded PDFs"""
    preview_data = []
    pdf_handler = PDFHandler()
    
    for pdf_file in pdf_files:
        try:
            text = pdf_handler.extract_text(pdf_file)
            amount = extract_primary_amount(text)
            
            if amount and amount > 1000:
                normalized = normalize_amount(amount)
                preview_data.append({
                    "PDF Name": pdf_file.name,
                    "Raw Amount": format_currency(amount),
                    "Normalized Amount": format_currency(normalized),
                    "Status": "✓"
                })
            else:
                # Try fallback to get any amount
                amounts = extract_amounts_from_text(text)
                valid_amounts = [amt for amt in amounts if amt > 1000]  # Filter small amounts
                if valid_amounts:
                    max_amount = max(valid_amounts)
                    normalized = normalize_amount(max_amount)
                    preview_data.append({
                        "PDF Name": pdf_file.name,
                        "Raw Amount": f"${max_amount:,.2f}",
                        "Normalized Amount": f"${normalized:,.2f}",
                        "Status": "⚠️"  # Warning because using fallback
                    })
                else:
                    preview_data.append({
                        "PDF Name": pdf_file.name,
                        "Raw Amount": "No valid amount found",
                        "Normalized Amount": "N/A",
                        "Status": "❌"
                    })
        except Exception as e:
            logger.error(f"Error processing PDF {pdf_file.name}: {str(e)}")
            preview_data.append({
                "PDF Name": pdf_file.name,
                "Raw Amount": "Error",
                "Normalized Amount": "Error",
                "Status": "❌"
            })
    
    return pd.DataFrame(preview_data)

def preview_match(selection: Dict, pdf_info: Dict):
    """Show side-by-side preview of selection and PDF for manual confirmation"""
    match_container = st.container()
    with match_container:
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("### Selection Details")
            st.write(f"ID: {selection['Selection ID']}")
            st.write(f"Amount: {selection['Selection Amount']}")
            for key, value in selection['Selection Data'].items():
                if key not in ['Selection ID', 'amount_column']:
                    st.write(f"{key}: {value}")

        with col2:
            st.write("### PDF Preview")
            st.write(f"Filename: {pdf_info['PDF Name']}")
            st.write(f"Found Amount: {pdf_info['PDF Amount']}")
            
            # Show PDF preview with error handling
            if pdf_info.get('Matched Pages'):
                for i, page_path in enumerate(pdf_info['Matched Pages']):
                    try:
                        img_path = Path(page_path)
                        if img_path.exists():
                            with Image.open(img_path) as img:
                                st.image(img, caption=f"Page {i+1}", use_column_width=True)
                        else:
                            st.warning(f"Image not found: {page_path}")
                    except Exception as e:
                        st.error(f"Error loading image: {str(e)}")
                        logger.error(f"Image load error: {str(e)}")

def interactive_matching(df: pd.DataFrame, pdf_files: List, unique_id_column: str, amount_column: str):
    """Interactive matching process with preview and confirmation"""
    try:
        # Initialize or get matching state
        if 'matching_state' not in st.session_state:
            matches = match_documents(
                df=df,
                unique_id_column=unique_id_column,
                amount_column=amount_column,
                pdf_files=pdf_files
            )
            st.session_state.matching_state = {
                'matches': matches,
                'current_index': 0,
                'confirmed_matches': [],
                'skipped_matches': [],
                'completed': False
            }
        
        state = st.session_state.matching_state
        matches = state['matches']
        confirmed_matches = state['confirmed_matches']
        
        # Group matches by Selection ID
        grouped_matches = {}
        for match in matches:
            selection_id = match['Selection ID']
            if selection_id not in grouped_matches:
                grouped_matches[selection_id] = []
            grouped_matches[selection_id].append(match)
        
        st.write("### Confirm Matches")
        st.write("Please review and confirm the matches for each selection.")
        
        for selection_id, potential_matches in grouped_matches.items():
            st.write(f"#### Selection {selection_id}")
            st.write(f"Amount: {potential_matches[0]['Selection Amount']}")
            
            if len(potential_matches) > 1:
                options = [f"{m['PDF Name']} - {m['PDF Amount']} ({m['Match Type']})" 
                          for m in potential_matches]
                selected = st.radio(
                    f"Choose the correct match for Selection {selection_id}:",
                    options,
                    key=f"selection_{selection_id}"
                )
                
                # Find the selected match
                selected_index = options.index(selected)
                confirmed_matches.append(potential_matches[selected_index])
            else:
                confirmed_matches.append(potential_matches[0])
        
        st.write("### Confirmed Matches")
        st.write(confirmed_matches)
        
    except Exception as e:
        st.error(f"Error during interactive matching: {str(e)}")
        logger.error(f"Interactive matching error: {str(e)}")

def _advance_match(state: Dict):
    """Helper to advance to next match"""
    state['current_index'] += 1
    if state['current_index'] >= state['total_matches']:
        state['completed'] = True
    st.rerun()  # Use st.rerun() instead of st.experimental_rerun()

def convert_pdf_to_images(pdf_path: Path) -> List[Path]:
    """Convert PDF pages to images and return list of image paths"""
    try:
        images = convert_from_path(pdf_path)
        image_paths = []
        for i, image in enumerate(images):
            image_path = pdf_path.with_name(f"{pdf_path.stem}_page_{i+1}.png")
            image.save(image_path, "PNG")
            image_paths.append(image_path)
        return image_paths
    except Exception as e:
        logger.error(f"Error converting PDF to images: {str(e)}")
        raise

def embed_images_in_excel(sheet, image_paths: List[Path], start_row: int):
    """Embed images into an Excel sheet starting from a specific row"""
    for i, image_path in enumerate(image_paths):
        img = ExcelImage(str(image_path))
        sheet.add_image(img, f'A{start_row + i}')
        # Optionally, remove the image file after embedding
        image_path.unlink(missing_ok=True)

def save_matches_to_excel_with_images(matches: List[Dict], output_path: Path, user_labels: Dict):
    """Save matches to an Excel file with embedded images"""
    wb = Workbook()
    summary_sheet = wb.active
    summary_sheet.title = "Summary"
    
    # Add report header
    summary_sheet['A1'] = "Audit Matcher Results"
    summary_sheet['A1'].font = openpyxl.styles.Font(bold=True, size=14)
    summary_sheet['A2'] = f"Generated: {time.strftime('%Y-%m-%d %H:%M:%S')}"
    
    # Add statistics
    row = 4
    summary_sheet[f'A{row}'] = "Statistics"
    summary_sheet[f'A{row}'].font = openpyxl.styles.Font(bold=True)
    row += 1
    
    for label, value in user_labels.items():
        summary_sheet[f'A{row}'] = label
        summary_sheet[f'B{row}'] = value
        row += 1
    
    for match in matches:
        sheet_name = str(match['Selection ID'])[:31]  # Excel sheet name limit
        sheet = wb.create_sheet(title=sheet_name)
        
        # Write selection data
        for col, (key, value) in enumerate(match['Selection Data'].items(), start=1):
            sheet.cell(row=1, column=col, value=f"{key}: {value}")
        
        # Add blank row for separation
        sheet.cell(row=2, column=1, value="")
        
        # Convert PDF to images and embed in Excel
        pdf_path = Path(match['PDF Path'])
        image_paths = convert_pdf_to_images(pdf_path)
        embed_images_in_excel(sheet, image_paths, start_row=3)
    
    # Remove default sheet created by openpyxl
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
    wb.save(output_path)
    logger.info(f"Excel file saved successfully to {output_path}")

def show_matching_summary(confirmed_matches: List[Dict], skipped_matches: List[Dict],
                        unique_id_column: str, amount_column: str) -> None:
    """Display summary of matching results"""
    st.write("### Matching Summary")
    
    # Show summary stats in columns
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"Total Confirmed: {len(confirmed_matches)}")
        if confirmed_matches:
            st.write("### Confirmed Matches")
            for match in confirmed_matches:
                with st.expander(f"Selection {match['Selection ID']} - {match['Selection Amount']}"):
                    st.write(f"Matched PDF: {match['PDF Name']}")
                    st.write(f"PDF Amount: {match['PDF Amount']}")
                    if match.get('Matched Pages'):
                        try:
                            for i, page_path in enumerate(match['Matched Pages']):
                                img_path = Path(page_path)
                                if img_path.exists():
                                    with Image.open(img_path) as img:
                                        st.image(img, caption=f"Page {i+1}", use_column_width=True)
                        except Exception as e:
                            st.error(f"Error loading PDF preview: {str(e)}")
    
    with col2:
        st.write(f"Total Skipped: {len(skipped_matches)}")
        if skipped_matches:
            for skip in skipped_matches:
                st.write(f"- Selection {skip['Selection ID']} ({skip['Selection Amount']})")
    
    # Export section
    if confirmed_matches:
        st.write("---")
        st.write("### Export Results")
        
        try:
            output_path = Path("matching_results_with_images.xlsx")
            user_labels = {
                "Unique ID Column": unique_id_column,
                "Amount Column": amount_column,
                "Processing Date": time.strftime("%Y-%m-%d %H:%M:%S"),
                "Total Matches": len(confirmed_matches),
                "Total Skipped": len(skipped_matches)
            }
            
            # Save Excel file with images using utility function
            save_matches_to_excel_with_images(confirmed_matches, output_path, user_labels)
            
            # Create download button with proper MIME type
            with open(output_path, "rb") as file:
                excel_data = file.read()
            
            st.download_button(
                label="📥 Download Results",
                data=excel_data,
                file_name="matching_results_with_images.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("✓ Excel report with images ready for download!")
            
            # Cleanup file after download button is created
            try:
                output_path.unlink(missing_ok=True)
            except Exception as e:
                logger.error(f"Error cleaning up Excel file: {e}")
                
        except Exception as e:
            logger.error(f"Error preparing Excel download: {e}")
            st.error(f"Error generating Excel file: {str(e)}")

def main():
    st.title("Audit Matcher")
    st.write("Welcome to Audit Matcher!")

    # Initialize session state if needed
    if 'app_state' not in st.session_state:
        st.session_state.app_state = {
            'matching_started': False,
            'df': None,
            'pdf_files': None,
            'columns': {'id': None, 'amount': None}
        }

    # File uploaders section
    st.write("### Step 1: Upload Excel File")
    selections_file = st.file_uploader(
        "Upload Selections Excel",
        type=['xlsx'],
        help="Select your audit selections Excel file"
    )

    if selections_file:
        try:
            # Load Excel file
            df = pd.read_excel(selections_file)
            if df.empty:
                st.error("The uploaded Excel file is empty")
                return

            st.success("✓ Excel file loaded successfully")
            st.session_state.app_state['df'] = df
            
            # Display raw data
            st.write("### Raw Data Preview:")
            st.dataframe(df.head())
            
            # Column selection
            columns = df.columns.tolist()
            unique_id_column = st.selectbox("Select ID Column", options=columns)
            amount_column = st.selectbox("Select Amount Column", options=columns)
            
            if amount_column:
                # Simple amount preview
                preview_df = pd.DataFrame({
                    'ID': df[unique_id_column],
                    'Amount': df[amount_column]
                })
                st.dataframe(preview_df)

            # PDF upload section
            st.write("### Step 2: Upload PDF Files")
            pdf_files = st.file_uploader(
                "Upload PDF Files",
                type=['pdf'],
                accept_multiple_files=True,
                help="Select one or more PDF files to match against the Excel data"
            )

            if pdf_files:
                st.session_state.app_state['pdf_files'] = pdf_files
                st.session_state.app_state['columns'] = {
                    'id': unique_id_column,
                    'amount': amount_column
                }
                
                st.success(f"✓ {len(pdf_files)} PDF file(s) uploaded successfully")
                
                # Show PDF previews
                with st.spinner("Extracting amounts from PDFs..."):
                    preview_df = preview_pdf_amounts(pdf_files)
                    st.write("### PDF Amount Preview")
                    st.dataframe(preview_df)

                # Start matching button
                if st.button("Start Matching") or st.session_state.app_state.get('matching_started'):
                    st.session_state.app_state['matching_started'] = True
                    interactive_matching(
                        df,
                        pdf_files,
                        unique_id_column,
                        amount_column
                    )

        except Exception as e:
            handle_error(f"Error: {str(e)}")
            logger.exception("Error in main:")
    else:
        st.info("Please upload an Excel file to begin")

    # Debug section
    if st.checkbox("Show Debug Info"):
        st.write("Session State:", st.session_state.app_state)

if __name__ == "__main__":
    main()
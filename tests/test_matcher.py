import unittest
from unittest.mock import patch, MagicMock
from decimal import Decimal
from matcher import extract_amounts_from_text, extract_primary_amount, match_documents
import pandas as pd
import os
from utils import save_matches_to_excel

class TestMatcherFunctions(unittest.TestCase):

    def test_extract_amounts_from_text(self):
        text = "The total amount is $1,234.56 and the other amount is $789.00."
        expected = [1234.56, 789.00]
        result = extract_amounts_from_text(text)
        self.assertEqual(result, expected)

    def test_extract_primary_amount(self):
        text = "The total amount is $1,234.56 and the other amount is $789.00."
        expected = 1234.56
        result = extract_primary_amount(text)
        self.assertEqual(result, expected)

    @patch('matcher.extract_text_from_pdf')
    @patch('matcher.extract_primary_amount')
    def test_match_documents(self, mock_extract_primary_amount, mock_extract_text_from_pdf):
        data = {
            'ID': [1, 2],
            'Amount': [1234.56, 789.00]
        }
        df = pd.DataFrame(data)
        pdf_files = [MagicMock(), MagicMock()]  # Mock PDF files

        mock_extract_text_from_pdf.return_value = "Mock text"
        mock_extract_primary_amount.return_value = 1234.56

        results = match_documents(df, 'ID', 'Amount', pdf_files)
        self.assertIsInstance(results, list)
        self.assertEqual(len(results), 2)

    def test_save_matches_to_excel(self):
        matches = [
            {
                'Selection ID': '1',
                'Selection Data': {'ID': '1', 'Amount': 1234.56},
                'Selection Amount': '$1,234.56',
                'PDF Name': 'test.pdf',
                'PDF Amount': '$1,234.56',
                'PDF Text': 'Mock text',
                'Match Type': 'Exact',
                'Match Score': 100,
                'Matched Pages': ['path/to/page1.png', 'path/to/page2.png']
            }
        ]
        user_labels = {
            "Unique ID Column": "ID",
            "Amount Column": "Amount",
            "Processing Date": "2023-10-01 12:00:00",
            "Total Matches": 1,
            "Total Skipped": 0
        }
        output_path = 'test_output.xlsx'
        save_matches_to_excel(matches, output_path, user_labels)
        self.assertTrue(os.path.exists(output_path))
        os.remove(output_path)

if __name__ == '__main__':
    unittest.main()
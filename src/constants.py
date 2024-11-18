
SUPPORT_CATEGORIES = [
    "Invoices",
    "Contracts",
    "Purchase Orders",
    "Payment Evidence",
    "Other Support"
]

OUTPUT_FOLDER = "audit_support_files"
PREVIEW_FILE = "support_preview.xlsx"
SUMMARY_FILE = "audit_summary.xlsx"

EXCEL_HEADERS = [
    "Selection ID",
    "Amount",
    "Description",
    *SUPPORT_CATEGORIES  # Unpack categories as additional headers
]
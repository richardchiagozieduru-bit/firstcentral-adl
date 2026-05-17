"""
Error Messages Module

Provides user-friendly error message mappings for technical errors.
Maps cryptic system errors to actionable, understandable messages.
"""

import logging
import re

logger = logging.getLogger(__name__)


# =============================================================================
# ERROR MESSAGE MAPPINGS
# Maps technical error patterns to user-friendly messages with suggested fixes
# =============================================================================

ERROR_MAPPINGS = {
    # -------------------------------------------------------------------------
    # VALIDATION ERRORS - Column/Data Issues
    # -------------------------------------------------------------------------
    'CUSTOMERID column is empty': {
        'title': 'Missing Customer IDs',
        'message': 'The Customer ID column in your file is empty or missing.',
        'fix': 'Please ensure every row has a unique Customer ID value in the appropriate column.',
        'icon': 'fa-id-card',
        'category': 'validation'
    },
    'ACCOUNTNUMBER column is empty': {
        'title': 'Missing Account Numbers',
        'message': 'The Account Number column is empty or missing.',
        'fix': 'Please add account numbers to all credit records in your file.',
        'icon': 'fa-hashtag',
        'category': 'validation'
    },
    'BVNNUMBER column is empty': {
        'title': 'Missing BVN Numbers',
        'message': 'Individual borrowers must have Bank Verification Numbers (BVN).',
        'fix': 'Add BVN numbers to the individual borrower sheet. Alternatively, provide NIN or Date of Birth as identifiers.',
        'icon': 'fa-fingerprint',
        'category': 'validation'
    },
    'SURNAME and FIRSTNAME columns are both empty': {
        'title': 'Missing Borrower Names',
        'message': 'Individual borrower records must have at least a surname or first name.',
        'fix': 'Please add borrower names (surname or first name) to the individual borrower sheet.',
        'icon': 'fa-user',
        'category': 'validation'
    },
    'Validation Failed': {
        'title': 'Data Validation Failed',
        'message': 'Your file contains data that does not meet the required format.',
        'fix': 'Please review the specific errors below and correct your file.',
        'icon': 'fa-exclamation-triangle',
        'category': 'validation'
    },
    
    # -------------------------------------------------------------------------
    # FILE ERRORS - File Format/Reading Issues  
    # -------------------------------------------------------------------------
    'file content does not match': {
        'title': 'Invalid File Format',
        'message': 'The uploaded file is not a valid Excel file.',
        'fix': 'Please ensure you are uploading a genuine Excel file (.xlsx or .xls). '
               'Files renamed with Excel extensions will not work.',
        'icon': 'fa-file-excel',
        'category': 'file'
    },
    "codec can't decode": {
        'title': 'File Encoding Error',
        'message': 'The uploaded text file contains characters in an unsupported encoding.',
        'fix': 'Open the file in a text editor (e.g. Notepad), then re-save it with '
               '"UTF-8" encoding selected. Alternatively, try saving as a new CSV from Excel.',
        'icon': 'fa-file-alt',
        'category': 'file'
    },
    'openpyxl': {
        'title': 'Excel File Error',
        'message': 'Unable to read the Excel file. The file may be corrupted or in an unsupported format.',
        'fix': 'Try opening the file in Excel, saving it as a new file, and uploading again.',
        'icon': 'fa-file-excel',
        'category': 'file'
    },
    'xlrd': {
        'title': 'Legacy Excel Error',
        'message': 'Unable to read the legacy Excel file (.xls format).',
        'fix': 'Try converting your file to the newer .xlsx format using Excel.',
        'icon': 'fa-file-excel',
        'category': 'file'
    },
    'No such file or directory': {
        'title': 'File Not Found',
        'message': 'The uploaded file could not be located for processing.',
        'fix': 'Please try uploading the file again. If the issue persists, contact support.',
        'icon': 'fa-question-circle',
        'category': 'file'
    },
    
    # -------------------------------------------------------------------------
    # PROCESSING ERRORS - Runtime Issues
    # -------------------------------------------------------------------------
    'JSON object must be str': {
        'title': 'Processing Data Error',
        'message': 'An internal data format error occurred during processing.',
        'fix': 'Please try uploading your file again. If the issue persists, contact support.',
        'icon': 'fa-cogs',
        'category': 'processing'
    },
    'NoneType': {
        'title': 'Missing Data Error',
        'message': 'Required data was missing during processing.',
        'fix': 'Ensure all required sheets and columns are present in your file.',
        'icon': 'fa-database',
        'category': 'processing'
    },
    'KeyError': {
        'title': 'Missing Column Error',
        'message': 'A required column was not found in your file.',
        'fix': 'Please verify that your file contains all required columns with correct headers.',
        'icon': 'fa-columns',
        'category': 'processing'
    },
    'MemoryError': {
        'title': 'File Too Large',
        'message': 'The file is too large to process in available memory.',
        'fix': 'Try splitting your file into smaller batches or contact support for assistance.',
        'icon': 'fa-memory',
        'category': 'processing'
    },
    'timeout': {
        'title': 'Processing Timeout',
        'message': 'The file took too long to process and timed out.',
        'fix': 'Large files may take longer. Try uploading during off-peak hours or splitting into smaller files.',
        'icon': 'fa-clock',
        'category': 'processing'
    },
    
    # -------------------------------------------------------------------------
    # SHEET ERRORS - Missing/Invalid Sheets
    # -------------------------------------------------------------------------
    'Individual Borrower': {
        'title': 'Individual Borrower Sheet Issue',
        'message': 'There was an issue with the Individual Borrower sheet.',
        'fix': 'Verify that the Individual Borrower sheet exists and contains valid data.',
        'icon': 'fa-user',
        'category': 'sheet'
    },
    'Corporate Borrower': {
        'title': 'Corporate Borrower Sheet Issue',
        'message': 'There was an issue with the Corporate Borrower sheet.',
        'fix': 'Verify that the Corporate Borrower sheet exists and contains valid data.',
        'icon': 'fa-building',
        'category': 'sheet'
    },
    'Credit Information': {
        'title': 'Credit Information Sheet Issue',
        'message': 'There was an issue with the Credit Information sheet.',
        'fix': 'Verify that the Credit Information sheet contains Account Numbers and Customer IDs.',
        'icon': 'fa-credit-card',
        'category': 'sheet'
    },
}


# Default fallback for unknown errors
DEFAULT_ERROR = {
    'title': 'Processing Error',
    'message': 'An unexpected error occurred during file processing.',
    'fix': 'Please try uploading again. If the problem persists, contact support with the error details below.',
    'icon': 'fa-exclamation-circle',
    'category': 'unknown'
}


def get_friendly_error(technical_error):
    """
    Convert a technical error message to a user-friendly format.
    
    Args:
        technical_error (str): The raw technical error message
        
    Returns:
        dict: User-friendly error with keys: title, message, fix, icon, category, technical_details
    """
    if not technical_error:
        return {**DEFAULT_ERROR, 'technical_details': 'No error details available'}
    
    # Convert to string if needed
    error_str = str(technical_error)
    
    # Search for matching patterns (case-insensitive partial match)
    for pattern, friendly_error in ERROR_MAPPINGS.items():
        if pattern.lower() in error_str.lower():
            logger.info(f"[ERROR MAPPING] Matched pattern '{pattern}' for error")
            return {
                **friendly_error,
                'technical_details': error_str
            }
    
    # No match found - return default with technical details
    logger.warning(f"[ERROR MAPPING] No pattern matched for error: {error_str[:100]}")
    return {
        **DEFAULT_ERROR,
        'technical_details': error_str
    }


def format_error_for_display(technical_error):
    """
    Format an error for HTML display with user-friendly message and technical details.
    
    Args:
        technical_error (str): The raw technical error message
        
    Returns:
        str: HTML-formatted error message
    """
    error_info = get_friendly_error(technical_error)
    
    html = f"""
    <div class="error-container">
        <div class="error-header">
            <i class="fas {error_info['icon']} me-2"></i>
            <strong>{error_info['title']}</strong>
        </div>
        <div class="error-message mt-2">
            {error_info['message']}
        </div>
        <div class="error-fix mt-2">
            <strong>Suggested Fix:</strong> {error_info['fix']}
        </div>
    </div>
    """
    
    return html.strip()


def get_user_friendly_message(technical_error):
    """
    Get a simple user-friendly message string (no HTML).
    
    Args:
        technical_error (str): The raw technical error message
        
    Returns:
        str: Simple user-friendly message
    """
    error_info = get_friendly_error(technical_error)
    return f"{error_info['title']}: {error_info['message']} {error_info['fix']}"


def parse_validation_errors(error_message):
    """
    Parse a validation error message with multiple sheet errors into structured format.
    
    Args:
        error_message (str): Multi-line validation error message
        
    Returns:
        list: List of dicts with sheet_name, errors, and friendly info
    """
    if not error_message:
        return []
    
    parsed_errors = []
    lines = error_message.split('\n')
    
    for line in lines:
        line = line.strip()
        if not line or line.startswith('❌') or line.startswith('Please correct'):
            continue
        
        # Try to parse "Sheet Name: error1. error2." format
        if ':' in line:
            parts = line.split(':', 1)
            sheet_name = parts[0].strip()
            errors = [e.strip() for e in parts[1].split('.') if e.strip()]
            
            parsed_errors.append({
                'sheet_name': sheet_name,
                'errors': errors,
                'friendly': get_friendly_error(line)
            })
    
    return parsed_errors

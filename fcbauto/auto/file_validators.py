"""
File Validators Module

Provides security validation for uploaded files using magic bytes detection
to prevent malicious file uploads disguised as Excel files.
"""

import logging

logger = logging.getLogger(__name__)


# Magic bytes signatures for Excel file types
# Reference: https://en.wikipedia.org/wiki/List_of_file_signatures
EXCEL_SIGNATURES = {
    # XLSX, XLSM, XLSB - ZIP-based Office Open XML format
    # Magic bytes: PK (50 4B 03 04)
    'zip_office': {
        'magic': b'\x50\x4B\x03\x04',
        'extensions': ['.xlsx', '.xlsm', '.xlsb'],
        'mime_types': [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel.sheet.macroEnabled.12',
            'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
        ],
    },
    # XLS - Legacy OLE2 compound document format
    # Magic bytes: D0 CF 11 E0 A1 B1 1A E1
    'ole2': {
        'magic': b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1',
        'extensions': ['.xls'],
        'mime_types': [
            'application/vnd.ms-excel',
        ],
    },
}

# Allowed extensions (XLSX, XLS, XLSB, XLSM, CSV, TXT)
ALLOWED_EXTENSIONS = {'.xlsx', '.xls', '.xlsb', '.xlsm', '.csv', '.txt'}

# User-friendly error messages
ERROR_MESSAGES = {
    'no_file': 'No file was uploaded. Please select a file and try again.',
    'empty_file': 'The uploaded file is empty. Please upload a valid Excel file.',
    'invalid_extension': 'Invalid file extension. Only Excel files (.xlsx, .xls, .xlsb, .xlsm) and delimited text files (.csv, .txt) are allowed.',
    'invalid_content': (
        'The file content does not match a valid Excel format. '
        'This may happen if a non-Excel file was renamed with an Excel extension. '
        'Please upload a genuine Excel file (.xlsx or .xls).'
    ),
    'file_read_error': 'Unable to read the uploaded file. Please try uploading again.',
}


def get_file_extension(filename):
    """
    Extract and normalize file extension from filename.
    
    Args:
        filename (str): Name of the file
        
    Returns:
        str: Lowercase file extension including the dot (e.g., '.xlsx')
    """
    if not filename:
        return ''
    
    import os
    _, ext = os.path.splitext(filename)
    return ext.lower()


def check_magic_bytes(file_buffer, expected_magic):
    """
    Check if file buffer starts with expected magic bytes.
    
    Args:
        file_buffer (bytes): First few bytes of the file
        expected_magic (bytes): Expected magic byte signature
        
    Returns:
        bool: True if magic bytes match
    """
    return file_buffer.startswith(expected_magic)


def validate_excel_file_type(uploaded_file):
    """
    Validate that an uploaded file is a genuine Excel file by checking
    both the file extension and magic bytes (file signature).
    
    This prevents attacks where malicious files are renamed with .xlsx/.xls
    extensions to bypass simple extension-based validation.
    
    Args:
        uploaded_file: Django UploadedFile object
        
    Returns:
        tuple: (is_valid: bool, error_message: str or None)
    """
    # Check if file exists
    if not uploaded_file:
        logger.warning("[FILE VALIDATION] No file provided")
        return False, ERROR_MESSAGES['no_file']
    
    # Check file size
    if uploaded_file.size == 0:
        logger.warning("[FILE VALIDATION] Empty file uploaded")
        return False, ERROR_MESSAGES['empty_file']
    
    # Get and validate extension
    filename = getattr(uploaded_file, 'name', '')
    extension = get_file_extension(filename)
    
    if extension not in ALLOWED_EXTENSIONS:
        logger.warning(f"[FILE VALIDATION] Invalid extension: {extension}")
        return False, ERROR_MESSAGES['invalid_extension']
    
    # Read magic bytes from file
    try:
        # Save current position
        current_position = uploaded_file.tell()
        
        # Read first 8 bytes for signature detection
        uploaded_file.seek(0)
        file_header = uploaded_file.read(8)
        
        # Reset file pointer to original position
        uploaded_file.seek(current_position)
        
    except Exception as e:
        logger.error(f"[FILE VALIDATION] Error reading file: {e}")
        return False, ERROR_MESSAGES['file_read_error']
    
    # Check magic bytes based on extension
    is_valid_content = False
    
    if extension in ['.xlsx', '.xlsm', '.xlsb']:
        # XLSX, XLSM, XLSB must be ZIP format (Office Open XML)
        is_valid_content = check_magic_bytes(file_header, EXCEL_SIGNATURES['zip_office']['magic'])
        
    elif extension == '.xls':
        # XLS must be OLE2 format
        is_valid_content = check_magic_bytes(file_header, EXCEL_SIGNATURES['ole2']['magic'])
    
    if not is_valid_content:
        logger.warning(
            f"[FILE VALIDATION] Magic bytes mismatch for {filename}. "
            f"Extension: {extension}, Header: {file_header[:4].hex()}"
        )
        return False, ERROR_MESSAGES['invalid_content']
    
    logger.info(f"[FILE VALIDATION] File validated successfully: {filename}")
    return True, None


def validate_csv_file(uploaded_file):
    """
    Validate that an uploaded CSV file is readable as plain text.

    CSV files have no magic bytes, so we just verify the extension and that
    the file content is decodable as UTF-8 text (not a renamed binary file).

    Args:
        uploaded_file: Django UploadedFile object

    Returns:
        tuple: (is_valid: bool, error_message: str or None)
    """
    if not uploaded_file:
        return False, ERROR_MESSAGES['no_file']

    if uploaded_file.size == 0:
        return False, ERROR_MESSAGES['empty_file']

    filename = getattr(uploaded_file, 'name', '')
    extension = get_file_extension(filename)

    if extension not in ('.csv', '.txt'):
        return False, ERROR_MESSAGES['invalid_extension']

    # Try reading the first 512 bytes to verify it is decodable plain text
    # (accept UTF-8, Windows-1252/CP1252, and Latin-1 encoded files)
    try:
        current_position = uploaded_file.tell()
        uploaded_file.seek(0)
        sample = uploaded_file.read(512)
        uploaded_file.seek(current_position)
        # Try common encodings; latin-1 is the guaranteed fallback that
        # decodes every byte, so if even that fails it's truly binary data.
        for enc in ('utf-8-sig', 'cp1252', 'latin-1'):
            try:
                sample.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            logger.warning(f"[FILE VALIDATION] {filename} could not be decoded as text")
            return False, "The file appears to be binary data, not a text/CSV file."
    except Exception as e:
        logger.error(f"[FILE VALIDATION] Error reading CSV file {filename}: {e}")
        return False, ERROR_MESSAGES['file_read_error']

    logger.info(f"[FILE VALIDATION] CSV file validated successfully: {filename}")
    return True, None


def get_detailed_validation_info(uploaded_file):
    """
    Get detailed validation information for debugging purposes.
    
    Args:
        uploaded_file: Django UploadedFile object
        
    Returns:
        dict: Detailed validation information
    """
    info = {
        'filename': getattr(uploaded_file, 'name', 'Unknown'),
        'size': getattr(uploaded_file, 'size', 0),
        'content_type': getattr(uploaded_file, 'content_type', 'Unknown'),
        'extension': '',
        'magic_bytes_hex': '',
        'detected_format': 'Unknown',
        'is_valid': False,
    }
    
    if uploaded_file:
        info['extension'] = get_file_extension(info['filename'])
        
        try:
            uploaded_file.seek(0)
            header = uploaded_file.read(8)
            uploaded_file.seek(0)
            
            info['magic_bytes_hex'] = header.hex()
            
            # Detect format from magic bytes
            if check_magic_bytes(header, EXCEL_SIGNATURES['zip_office']['magic']):
                info['detected_format'] = 'ZIP/Office Open XML (XLSX/XLSM/XLSB)'
            elif check_magic_bytes(header, EXCEL_SIGNATURES['ole2']['magic']):
                info['detected_format'] = 'OLE2 Compound Document (XLS)'
            else:
                info['detected_format'] = 'Unknown/Invalid'
                
            is_valid, _ = validate_excel_file_type(uploaded_file)
            info['is_valid'] = is_valid
            
        except Exception as e:
            info['error'] = str(e)
    
    return info

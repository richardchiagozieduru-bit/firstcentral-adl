"""
Custom exceptions for the Credit Bureau Automation system.

These exceptions provide more specific error handling and make debugging easier
by including relevant context (sheet names, file paths, etc.) with each error.

Usage:
    from .exceptions import DataValidationError, FileProcessingError

    # Raising
    raise DataValidationError("CUSTOMERID is empty", sheet_name="Individual Borrower")

    # Catching
    try:
        process_file(...)
    except DataValidationError as e:
        logger.warning(f"Validation failed on {e.sheet_name}: {e}")
"""


class CreditBureauException(Exception):
    """
    Base exception for all Credit Bureau application errors.
    
    All custom exceptions inherit from this, allowing you to catch
    any app-specific error with: except CreditBureauException
    """
    pass


class DataValidationError(CreditBureauException):
    """
    Raised when uploaded data fails validation rules.
    
    Examples:
        - Required column is empty (CUSTOMERID, ACCOUNTNUMBER)
        - Missing required sheets
        - Invalid data formats
    
    Attributes:
        sheet_name: Name of the sheet where validation failed
        column: Name of the column that failed validation
        row_count: Number of affected rows (if applicable)
    """
    def __init__(self, message, sheet_name=None, column=None, row_count=None):
        self.sheet_name = sheet_name
        self.column = column
        self.row_count = row_count
        super().__init__(message)
    
    def __str__(self):
        parts = [super().__str__()]
        if self.sheet_name:
            parts.append(f"Sheet: {self.sheet_name}")
        if self.column:
            parts.append(f"Column: {self.column}")
        if self.row_count:
            parts.append(f"Affected rows: {self.row_count}")
        return " | ".join(parts)


class FileProcessingError(CreditBureauException):
    """
    Raised when file processing operations fail.
    
    Examples:
        - Excel file cannot be read
        - File is corrupted or invalid format
        - Sheet not found in workbook
    
    Attributes:
        file_path: Path to the file that caused the error
        sheet_name: Name of the sheet (if applicable)
    """
    def __init__(self, message, file_path=None, sheet_name=None):
        self.file_path = file_path
        self.sheet_name = sheet_name
        super().__init__(message)
    
    def __str__(self):
        parts = [super().__str__()]
        if self.file_path:
            parts.append(f"File: {self.file_path}")
        if self.sheet_name:
            parts.append(f"Sheet: {self.sheet_name}")
        return " | ".join(parts)


class MergeError(CreditBureauException):
    """
    Raised when borrower-credit record merging fails.
    
    Examples:
        - No matching CUSTOMERID found
        - Merge produced unexpected results
        - Empty merge result
    
    Attributes:
        borrower_type: 'individual' or 'corporate'
        unmatched_count: Number of records that couldn't be matched
    """
    def __init__(self, message, borrower_type=None, unmatched_count=None):
        self.borrower_type = borrower_type
        self.unmatched_count = unmatched_count
        super().__init__(message)


class OutputGenerationError(CreditBureauException):
    """
    Raised when output file generation fails.
    
    Examples:
        - Cannot write Excel file
        - TXT file generation fails
        - Disk space issues
    
    Attributes:
        output_type: 'xlsx' or 'txt'
        file_path: Path where output was being written
    """
    def __init__(self, message, output_type=None, file_path=None):
        self.output_type = output_type
        self.file_path = file_path
        super().__init__(message)


class VerificationError(CreditBureauException):
    """
    Raised when human verification process encounters issues.
    
    Examples:
        - Verification data not found
        - Invalid verification state
        - Session expired during verification
    
    Attributes:
        session_id: Upload session ID
    """
    def __init__(self, message, session_id=None):
        self.session_id = session_id
        super().__init__(message)

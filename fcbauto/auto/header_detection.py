"""
Header Row Auto-Detection Utility

This module provides functionality to automatically detect the header row
in Excel worksheets, handling cases where the first row contains title
information or notes rather than actual column headers.
"""

import logging
import pandas as pd
from rapidfuzz import fuzz

logger = logging.getLogger(__name__)

# Known column names across all sheet types for header detection
# These are the most distinctive/common columns that help identify headers
KNOWN_COLUMNS = [
    # Individual Borrower
    'CUSTOMERID', 'SURNAME', 'FIRSTNAME', 'MIDDLENAME', 'DATEOFBIRTH',
    'BVNNUMBER', 'NATIONALIDENTITYNUMBER', 'PASSPORTNUMBER', 'GENDER',
    'MOBILENUMBER', 'EMAILADDRESS', 'PRIMARYADDRESSLINE1', 'MARITALSTATUS',
    
    # Corporate Borrower
    'BUSINESSNAME', 'BUSINESSREGISTRATIONNUMBER', 'DATEOFINCORPORATION',
    'CUSTOMERBRANCHCODE', 'BUSINESSOFFICEADDRESSLINE1', 'TAXID',
    
    # Credit Information
    'ACCOUNTNUMBER', 'ACCOUNTSTATUS', 'ACCOUNTSTATUSDATE', 'LOANEFFECTIVEDATE',
    'CREDITLIMIT', 'AVAILEDLIMIT', 'OUTSTANDINGBALANCE', 'DAYSINARREARS',
    'FACILITYTYPE', 'FACILITYTENOR', 'MATURITYDATE', 'LOANCLASSIFICATION',
    'CURRENCY', 'REPAYMENTFREQUENCY', 'COLLATERALTYPE',
    
    # Principal Officers
    'PRINCIPALOFFICER1SURNAME', 'PRINCIPALOFFICER1FIRSTNAME', 'POSITIONINBUSINESS',
    
    # Guarantors
    'INDIVIDUALGUARANTORSURNAME', 'GUARANTORDATEOFBIRTHINCORPORATION', 'TYPEOFGUARANTEE',
    
    # Common across types
    'BRANCHCODE', 'DATA'
]


def score_row_as_header(row_values, min_score=70):
    """
    Calculate a header confidence score for a row.
    
    Args:
        row_values: List of cell values from a row
        min_score: Minimum fuzzy match score to count as a match (default 70)
    
    Returns:
        int: Number of cells that match known column names
    """
    if not row_values:
        return 0
    
    matches = 0
    for value in row_values:
        if value is None:
            continue
        
        # Clean and normalize the value
        clean_value = str(value).upper().strip()
        clean_value = ''.join(c for c in clean_value if c.isalnum())
        
        if not clean_value:
            continue
        
        # Check fuzzy match against known columns
        for known_col in KNOWN_COLUMNS:
            score = fuzz.ratio(clean_value, known_col)
            if score >= min_score:
                matches += 1
                break  # Count each cell only once
    
    return matches


def find_header_row(file_path, sheet_name, max_scan_rows=10):
    """
    Find the row containing actual column headers by fuzzy matching.
    
    Scans the first N rows of a worksheet and identifies which row
    most likely contains the column headers based on fuzzy matching
    against known column names.
    
    Args:
        file_path: Path to the Excel file
        sheet_name: Name of the worksheet to analyze
        max_scan_rows: Maximum number of rows to scan (default 10)
    
    Returns:
        int: 0-indexed row number containing headers (default 0 if no match found)
    """
    try:
        # Determine engine based on file extension (.xlsb requires pyxlsb)
        engine = 'pyxlsb' if file_path.lower().endswith('.xlsb') else None
        
        # Read first N rows without assuming any header
        raw_df = pd.read_excel(
            file_path, 
            sheet_name=sheet_name, 
            header=None, 
            nrows=max_scan_rows,
            dtype=object,
            engine=engine
        )
        
        if raw_df.empty:
            logger.info(f"[HEADER DETECTION] Sheet '{sheet_name}' is empty, defaulting to row 0")
            return 0
        
        best_row = 0
        best_score = 0
        
        for row_idx in range(len(raw_df)):
            row_values = raw_df.iloc[row_idx].tolist()
            score = score_row_as_header(row_values)
            
            logger.debug(f"[HEADER DETECTION] Row {row_idx} score: {score}")
            
            if score > best_score:
                best_score = score
                best_row = row_idx
        
        # Only use detected row if we have a meaningful match (at least 2 columns matched)
        if best_score >= 2:
            logger.info(f"[HEADER DETECTION] Sheet '{sheet_name}': detected headers at row {best_row} (score: {best_score})")
            return best_row
        else:
            logger.info(f"[HEADER DETECTION] Sheet '{sheet_name}': no strong header match found, using row 0")
            return 0
            
    except Exception as e:
        logger.warning(f"[HEADER DETECTION] Error scanning sheet '{sheet_name}': {e}. Defaulting to row 0")
        return 0

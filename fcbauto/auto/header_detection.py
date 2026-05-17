import logging
import csv
import pandas as pd
from rapidfuzz import fuzz

logger = logging.getLogger(__name__)


def detect_csv_encoding(file_path):
    """
    Detect the character encoding of a CSV/TXT file.

    Tries encodings in order of preference:
      1. utf-8-sig  (UTF-8 with optional BOM)
      2. cp1252     (Windows-1252 — common in Nigerian/Windows-generated files)
      3. latin-1    (ISO-8859-1 — decodes every byte, guaranteed fallback)

    Args:
        file_path: Path to the file

    Returns:
        str: The first encoding that successfully reads the file
    """
    for enc in ('utf-8-sig', 'cp1252', 'latin-1'):
        try:
            with open(file_path, 'r', encoding=enc) as f:
                # Read the ENTIRE file to catch non-UTF-8 bytes that may
                # appear far into the file (e.g. Windows-1252 characters
                # at position 40k+).  A small sample (8 KB) is not enough
                # because the first chunk may be valid UTF-8 while later
                # content contains cp1252-only bytes like 0x92 (right quote).
                f.read()
            logger.debug(f"[ENCODING DETECT] '{file_path}' detected as {enc}")
            return enc
        except UnicodeDecodeError:
            continue
    return 'latin-1'  # latin-1 never raises — truly universal fallback


def detect_csv_delimiter(file_path, encoding=None):
    """
    Detect the column delimiter of a CSV/TXT file.

    Uses Python's csv.Sniffer on a sample of the file.  Falls back to tab
    for .txt files and comma for .csv files when the sniffer cannot decide.

    Args:
        file_path: Path to the file
        encoding:  Character encoding (if already known).  If None the
                   function calls detect_csv_encoding() internally.

    Returns:
        str: The detected delimiter character (e.g. '\\t', ',', '|', ';')
    """
    if encoding is None:
        encoding = detect_csv_encoding(file_path)

    try:
        with open(file_path, 'r', encoding=encoding, errors='replace') as f:
            sample = f.read(16384)  # 16 KB sample is plenty for sniffing

        dialect = csv.Sniffer().sniff(sample, delimiters='\t,|;')
        detected = dialect.delimiter
        logger.info(f"[DELIMITER DETECT] '{file_path}' delimiter detected as {repr(detected)}")
        return detected

    except csv.Error:
        # Sniffer could not determine a delimiter — use a sensible default
        # based on the file extension.
        if file_path.lower().endswith('.txt'):
            logger.info(f"[DELIMITER DETECT] Sniffer failed for '{file_path}', defaulting to tab (TXT file)")
            return '\t'
        else:
            logger.info(f"[DELIMITER DETECT] Sniffer failed for '{file_path}', defaulting to comma (CSV file)")
            return ','


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


def find_header_row_csv(file_path, max_scan_rows=10):
    """
    Find the header row in a CSV file by fuzzy matching column names.

    Works the same as find_header_row() but reads the file as CSV instead
    of an Excel worksheet.

    Args:
        file_path: Path to the CSV file
        max_scan_rows: Maximum number of rows to scan (default 10)

    Returns:
        int: 0-indexed row number containing headers (default 0 if no match found)
    """
    try:
        encoding = detect_csv_encoding(file_path)
        sep = detect_csv_delimiter(file_path, encoding=encoding)
        raw_df = pd.read_csv(
            file_path,
            header=None,
            nrows=max_scan_rows,
            dtype=object,
            sep=sep,
            encoding=encoding,
            encoding_errors='replace',
        )

        if raw_df.empty:
            logger.info(f"[CSV HEADER DETECTION] File '{file_path}' is empty, defaulting to row 0")
            return 0

        best_row = 0
        best_score = 0

        for row_idx in range(len(raw_df)):
            row_values = raw_df.iloc[row_idx].tolist()
            score = score_row_as_header(row_values)

            logger.debug(f"[CSV HEADER DETECTION] Row {row_idx} score: {score}")

            if score > best_score:
                best_score = score
                best_row = row_idx

        if best_score >= 2:
            logger.info(
                f"[CSV HEADER DETECTION] Detected headers at row {best_row} "
                f"(score: {best_score}) in '{file_path}'"
            )
            return best_row
        else:
            logger.info(f"[CSV HEADER DETECTION] No strong header match found in '{file_path}', using row 0")
            return 0

    except Exception as e:
        logger.warning(f"[CSV HEADER DETECTION] Error scanning '{file_path}': {e}. Defaulting to row 0")
        return 0


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

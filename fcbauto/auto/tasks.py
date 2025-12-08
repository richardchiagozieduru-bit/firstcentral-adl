import os
import glob
import shutil
import json
import time
import gc
import pandas as pd
import numpy as np
import openpyxl
from datetime import datetime
from django.conf import settings
from django.core.files.storage import FileSystemStorage
from datetime import datetime as dt
from .models import UploadSession
from .map import (
    consu_mapping, comm_mapping, credit_mapping, guar_mapping, prin_mapping,
    consumer_merged_mapping, commercial_merged_mapping
)
from .views import (
    # Import all the processing functions
    ensure_all_sheets_exist, clean_sheet_name, convert_numpy,
    remove_special_characters, make_column_names_unique,
    preprocess_tenor_from_headers, rename_columns_with_fuzzy_rapidfuzz,
    process_dates, process_names, replace_ampersands, process_special_characters,
    process_nationality, process_gender, process_states, process_marital_status,
    process_borrower_type, process_employment_status, process_phone_columns,
    process_title, process_account_status, process_loan_type, process_currency,
    process_repayment, process_classification, process_collateral_type,
    process_loan_tenor, clear_previous_info_columns, process_numeric_columns,
    fill_data_column, fill_depend_column, process_identity_numbers,
    process_passport_number, process_business_id, process_bvn_number,
    process_DriversLicense, process_otherid, process_tax_numbers,
    process_collateral_details, positioninBusiness, trim_strings_to_59,
    remove_duplicates, merge_individual_borrowers, merge_corporate_borrowers,
    split_commercial_entities, split_consumer_entities, reorder_commercial_columns,
    reorder_consumer_columns, extract_date_from_filename
)


# Constants for chunked processing
LARGE_SHEET_THRESHOLD = 50000  # Process sheets with >50k rows in chunks
CHUNK_SIZE = 50000  # Rows per chunk for large sheets


# ============================================================================
# COLUMN VALIDATION FUNCTIONS (Data Quality Checks)
# ============================================================================

def is_column_populated(df, column_name):
    """
    Check if a column has at least one non-empty value in the entire dataset.
    Uses optimized pandas operations with early termination for performance.
    
    Args:
        df (pd.DataFrame): DataFrame to check
        column_name (str): Name of column to validate
    
    Returns:
        bool: True if column has at least one non-empty value, False otherwise
    """
    # Check if column exists
    if column_name not in df.columns:
        return False
    
    # Quick null check first (early exit if all null)
    if not df[column_name].notna().any():
        return False
    
    # Drop null values for efficiency
    non_null_values = df[column_name].dropna()
    if len(non_null_values) == 0:
        return False
    
    # Check for non-empty strings (early exit on first match)
    return (non_null_values.astype(str).str.strip() != '').any()


def validate_required_columns(processed_sheets):
    """
    Validate required columns after fuzzy mapping/column renaming.
    Checks business rules for each sheet type.
    
    Validation Rules:
    - Credit Information: CUSTOMERID + ACCOUNTNUMBER both required
    - Consumer/Commercial Merged: ACCOUNTNUMBER required
    - Individual Borrower: CUSTOMERID required + (SURNAME or FIRSTNAME) required
    - Corporate Borrower: CUSTOMERID required
    - Principal Officers: CUSTOMERID required
    - Guarantors Information: Skip validation
    - Auto-generated (empty): Skip validation
    
    Args:
        processed_sheets (dict): Dictionary of sheet_name -> DataFrame
    
    Returns:
        tuple: (is_valid: bool, error_messages: list of str)
    """
    error_messages = []
    
    for sheet_key, df in processed_sheets.items():
        # Skip auto-generated empty sheets (headers only, no data)
        if df.empty:
            print(f"[VALIDATION] Skipping auto-generated empty sheet: {sheet_key}")
            continue
        
        # Skip guarantors information
        if sheet_key == 'guarantorsinformation':
            print(f"[VALIDATION] Skipping guarantors sheet: {sheet_key}")
            continue
        
        sheet_errors = []
        
        # Rule 1: Consumer/Commercial Merged - ACCOUNTNUMBER required
        if sheet_key in ['consumermerged', 'commercialmerged']:
            if not is_column_populated(df, 'ACCOUNTNUMBER'):
                sheet_errors.append("ACCOUNTNUMBER column is empty")
        
        # Rule 2: Credit Information - BOTH CUSTOMERID and ACCOUNTNUMBER required
        elif sheet_key == 'creditinformation':
            if not is_column_populated(df, 'CUSTOMERID'):
                sheet_errors.append("CUSTOMERID column is empty")
            if not is_column_populated(df, 'ACCOUNTNUMBER'):
                sheet_errors.append("ACCOUNTNUMBER column is empty")
        
        # Rule 3: Individual Borrower - CUSTOMERID + (SURNAME or FIRSTNAME) + BVNNUMBER required
        elif sheet_key == 'individualborrowertemplate':
            if not is_column_populated(df, 'CUSTOMERID'):
                sheet_errors.append("CUSTOMERID column is empty")
            
            has_surname = is_column_populated(df, 'SURNAME')
            has_firstname = is_column_populated(df, 'FIRSTNAME')
            
            if not has_surname and not has_firstname:
                sheet_errors.append("SURNAME and FIRSTNAME columns are both empty (at least one required)")
            
            if not is_column_populated(df, 'BVNNUMBER'):
                sheet_errors.append("BVNNUMBER column is empty")
        
        # Rule 4: Corporate Borrower - CUSTOMERID required
        elif sheet_key == 'corporateborrowertemplate':
            if not is_column_populated(df, 'CUSTOMERID'):
                sheet_errors.append("CUSTOMERID column is empty")
        
        # Rule 5: Principal Officers - CUSTOMERID required
        elif sheet_key == 'principalofficerstemplate':
            if not is_column_populated(df, 'CUSTOMERID'):
                sheet_errors.append("CUSTOMERID column is empty")
        
        # Add sheet errors to master list
        if sheet_errors:
            # Format sheet name for display
            sheet_display_name = sheet_key.replace('template', ' Template').replace('information', ' Information').title()
            # Format: "Sheet Name: error1. error2."
            formatted_errors = ". ".join(sheet_errors) + "."
            error_messages.append(f"{sheet_display_name}: {formatted_errors}")
    
    # Return validation result
    if error_messages:
        full_error_message = (
            "❌ Validation Failed - Required Columns Are Empty.\n"
            + "\n".join(error_messages)
            + "\nPlease correct these issues and re-upload your file."
        )
        return False, full_error_message
    
    return True, "✓ All required columns validated successfully"


# ============================================================================
# PARQUET TEMP STORAGE FUNCTIONS (Memory Management for Large Files)
# ============================================================================

def get_temp_parquet_dir(session_id):
    """
    Get the temporary Parquet directory path for a session
    
    Args:
        session_id: UploadSession ID
    
    Returns:
        Path to temp directory
    """
    temp_dir = os.path.join(settings.MEDIA_ROOT, 'temp', str(session_id))
    os.makedirs(temp_dir, exist_ok=True)
    return temp_dir


def save_to_parquet(df, session_id, sheet_name):
    """
    Save DataFrame to temporary Parquet file
    
    Args:
        df: DataFrame to save (can be empty but must have proper columns)
        session_id: UploadSession ID
        sheet_name: Name identifier for the sheet (e.g., 'consumer', 'credit', 'corporate')
    
    Returns:
        Path to saved Parquet file
    """
    temp_dir = get_temp_parquet_dir(session_id)
    parquet_path = os.path.join(temp_dir, f'{sheet_name}_processed.parquet')
    
    # Log saving status - important for empty auto-generated sheets
    row_status = f"({len(df)} rows)" if not df.empty else "(empty but with headers)"
    print(f"[PARQUET] Saving {sheet_name} to {parquet_path} {row_status}")
    
    df.to_parquet(parquet_path, engine='pyarrow', compression='snappy', index=False)
    print(f"[PARQUET] Successfully saved {sheet_name}")
    
    return parquet_path


def load_from_parquet(session_id, sheet_name, columns=None):
    """
    Load DataFrame from temporary Parquet file
    
    Args:
        session_id: UploadSession ID
        sheet_name: Name identifier for the sheet
        columns: Optional list of specific columns to load (memory optimization)
    
    Returns:
        DataFrame loaded from Parquet
    """
    temp_dir = get_temp_parquet_dir(session_id)
    parquet_path = os.path.join(temp_dir, f'{sheet_name}_processed.parquet')
    
    if not os.path.exists(parquet_path):
        raise FileNotFoundError(f"Parquet file not found: {parquet_path}")
    
    print(f"[PARQUET] Loading {sheet_name} from {parquet_path}" + 
          (f" (columns: {columns})" if columns else ""))
    
    df = pd.read_parquet(parquet_path, engine='pyarrow', columns=columns)
    print(f"[PARQUET] Successfully loaded {sheet_name} ({len(df)} rows)")
    
    return df


def cleanup_temp_files(session_id, keep_unmatched_credits=True):
    """
    Delete temporary Parquet files for a session, optionally keeping unmatched_credits for reporting
    
    Args:
        session_id: UploadSession ID
        keep_unmatched_credits: If True, preserve unmatched_credits.parquet for Excel report generation
    """

    
    temp_dir = get_temp_parquet_dir(session_id)
    
    if os.path.exists(temp_dir):
        if keep_unmatched_credits:
            # Delete individual files except unmatched_credits
            files_to_delete = glob.glob(os.path.join(temp_dir, '*_processed.parquet'))
            unmatched_file = os.path.normpath(os.path.join(temp_dir, 'unmatched_credits_processed.parquet'))
            
            print(f"[PARQUET] Cleanup with keep_unmatched_credits=True")
            print(f"[PARQUET] Unmatched file to preserve: {unmatched_file}")
            print(f"[PARQUET] Files found: {len(files_to_delete)}")
            
            deleted_count = 0
            preserved = False
            for file_path in files_to_delete:
                normalized_path = os.path.normpath(file_path)
                if normalized_path == unmatched_file:
                    print(f"[PARQUET] ✓ Preserving: {file_path}")
                    preserved = True
                else:
                    try:
                        os.remove(file_path)
                        deleted_count += 1
                        print(f"[PARQUET] ✗ Deleted: {file_path}")
                    except Exception as e:
                        print(f"[PARQUET] Error deleting {file_path}: {e}")
            
            print(f"[PARQUET] Cleanup complete: deleted {deleted_count} files, preserved unmatched_credits: {preserved}")
        else:
            # Delete entire directory
            print(f"[PARQUET] Cleaning up entire temp directory: {temp_dir}")
            shutil.rmtree(temp_dir)
            print(f"[PARQUET] Cleanup complete")
    else:
        print(f"[PARQUET] No temp directory to cleanup for session {session_id}")


# ============================================================================
# SHEET PROCESSING FUNCTIONS
# ============================================================================

def process_single_sheet(sheet_data, cleaned_name):
    """
    Apply all data transformation functions to a single sheet
    
    Args:
        sheet_data: DataFrame containing sheet data
        cleaned_name: Cleaned sheet name for mapping detection
    
    Returns:
        Processed DataFrame
    """
    cleaned_df = sheet_data.copy()
    
    # Comprehensive null value cleaning
    null_values = [
        'N/A', 'N.A', 'None', 'NaN', 'null', 'n/a', '#N/A', 'NIL', 'Nill', 'NA',
        'NULL', 'Null', 'NONE', 'UNKNOWN', 'unknown', 'Unknown', 'BLANK', 'blank', 'Blank',
        'nan', 'NAN', 'nil', 'Nil',
        'N.A.', 'n.a.', 'NOT AVAILABLE', 'not available', 'Not Available',
        'NOT APPLICABLE', 'not applicable', 'Not Applicable', 'MISSING', 'missing', 'Missing'
    ]
    cleaned_df.replace(null_values, '', inplace=True)
    
    # Clean column names
    cleaned_df.columns = [str(col).upper().strip() for col in cleaned_df.columns]
    cleaned_df = preprocess_tenor_from_headers(cleaned_df)
    cleaned_df.columns = [remove_special_characters(col) for col in cleaned_df.columns]
    cleaned_df = make_column_names_unique(cleaned_df)
    cleaned_df.columns = [remove_special_characters(col) for col in cleaned_df.columns]
    
    # Fuzzy column mapping based on sheet type
    if cleaned_name == 'individualborrowertemplate':
        cleaned_df = rename_columns_with_fuzzy_rapidfuzz(cleaned_df, consu_mapping)
    elif cleaned_name == 'corporateborrowertemplate':
        cleaned_df = rename_columns_with_fuzzy_rapidfuzz(cleaned_df, comm_mapping)
    elif cleaned_name == 'principalofficerstemplate':
        cleaned_df = rename_columns_with_fuzzy_rapidfuzz(cleaned_df, prin_mapping)
    elif cleaned_name == 'creditinformation':
        cleaned_df = rename_columns_with_fuzzy_rapidfuzz(cleaned_df, credit_mapping)
    elif cleaned_name == 'guarantorsinformation':
        cleaned_df = rename_columns_with_fuzzy_rapidfuzz(cleaned_df, guar_mapping)
    elif cleaned_name == 'consumermerged':
        cleaned_df = rename_columns_with_fuzzy_rapidfuzz(cleaned_df, consumer_merged_mapping)
    elif cleaned_name == 'commercialmerged':
        cleaned_df = rename_columns_with_fuzzy_rapidfuzz(cleaned_df, commercial_merged_mapping)
    
    # Apply all data cleaning transformations
    cleaned_df = process_dates(cleaned_df)
    cleaned_df = process_names(cleaned_df)
    cleaned_df = replace_ampersands(cleaned_df)
    cleaned_df = process_special_characters(cleaned_df)
    cleaned_df = process_nationality(cleaned_df)
    cleaned_df = process_gender(cleaned_df)
    cleaned_df = process_states(cleaned_df)
    cleaned_df = process_marital_status(cleaned_df)
    cleaned_df = process_borrower_type(cleaned_df)
    cleaned_df = process_employment_status(cleaned_df)
    cleaned_df = process_phone_columns(cleaned_df)
    cleaned_df = process_title(cleaned_df)
    cleaned_df = process_account_status(cleaned_df)
    cleaned_df = process_loan_type(cleaned_df)
    cleaned_df = process_currency(cleaned_df)
    cleaned_df = process_repayment(cleaned_df)
    cleaned_df = process_classification(cleaned_df)
    cleaned_df = process_collateral_type(cleaned_df)
    cleaned_df = process_loan_tenor(cleaned_df)
    cleaned_df = clear_previous_info_columns(cleaned_df)
    cleaned_df = process_numeric_columns(cleaned_df)
    cleaned_df = fill_data_column(cleaned_df)
    cleaned_df = fill_depend_column(cleaned_df)
    cleaned_df = process_identity_numbers(cleaned_df)
    cleaned_df = process_passport_number(cleaned_df)
    cleaned_df = process_business_id(cleaned_df)
    cleaned_df = process_bvn_number(cleaned_df)
    cleaned_df = process_DriversLicense(cleaned_df)
    cleaned_df = process_otherid(cleaned_df)
    cleaned_df = process_tax_numbers(cleaned_df)
    cleaned_df = process_collateral_details(cleaned_df)
    cleaned_df = positioninBusiness(cleaned_df)
    cleaned_df = trim_strings_to_59(cleaned_df)
    cleaned_df = remove_duplicates(cleaned_df)
    
    return cleaned_df


def process_large_sheet_chunked(file_path, sheet_name, cleaned_name, chunk_size=CHUNK_SIZE):
    """
    Process large sheets (>50k rows) in chunks to avoid memory issues and timeouts.
    Uses openpyxl for streaming read to minimize memory footprint.
    
    Args:
        file_path: Path to Excel file
        sheet_name: Name of sheet to process
        cleaned_name: Cleaned sheet name for mapping detection
        chunk_size: Number of rows per chunk (default 50,000)
    
    Returns:
        Fully processed DataFrame
    """
    print(f"[CHUNKED PROCESSING] Processing large sheet '{sheet_name}' in chunks of {chunk_size} rows")
    
    processed_chunks = []
    
    # Open workbook in read-only mode for memory efficiency
    workbook = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    sheet = workbook[sheet_name]
    
    # Get header row from first row
    header = [cell.value for cell in sheet[1]]
    print(f"[CHUNKED PROCESSING] Header extracted: {len(header)} columns")
    
    chunk_data = []
    chunk_number = 0
    
    # Iterate through rows starting from row 2 (skip header)
    for row_index, row in enumerate(sheet.iter_rows(min_row=2), start=1):
        row_values = [cell.value for cell in row]
        chunk_data.append(row_values)
        
        # When chunk is full, process it
        if row_index % chunk_size == 0:
            chunk_number += 1
            print(f"[CHUNKED PROCESSING] Processing chunk {chunk_number} ({len(chunk_data)} rows)")
            
            # Create DataFrame from chunk
            chunk_df = pd.DataFrame(chunk_data, columns=header)
            
            # Convert all columns to string type
            for col in chunk_df.columns:
                chunk_df[col] = chunk_df[col].astype(str)
                chunk_df[col] = chunk_df[col].replace({'nan': '', 'None': '', 'NaN': ''})
            
            # Process this chunk
            processed_chunk = process_single_sheet(chunk_df, cleaned_name)
            processed_chunks.append(processed_chunk)
            
            # Clear chunk data and force garbage collection
            chunk_data = []
            del chunk_df
            gc.collect()
    
    # Process the final chunk if any rows remain
    if chunk_data:
        chunk_number += 1
        print(f"[CHUNKED PROCESSING] Processing final chunk {chunk_number} ({len(chunk_data)} rows)")
        
        chunk_df = pd.DataFrame(chunk_data, columns=header)
        
        # Convert all columns to string type
        for col in chunk_df.columns:
            chunk_df[col] = chunk_df[col].astype(str)
            chunk_df[col] = chunk_df[col].replace({'nan': '', 'None': '', 'NaN': ''})
        
        processed_chunk = process_single_sheet(chunk_df, cleaned_name)
        processed_chunks.append(processed_chunk)
        
        del chunk_df
        gc.collect()
    
    # Close workbook
    workbook.close()
    
    # Combine all processed chunks
    if not processed_chunks:
        print(f"[CHUNKED PROCESSING] Warning: No data chunks processed for sheet '{sheet_name}'")
        return pd.DataFrame(columns=header)
    
    print(f"[CHUNKED PROCESSING] Combining {len(processed_chunks)} processed chunks")
    final_result = pd.concat(processed_chunks, ignore_index=True)
    
    # Clean up chunk list
    del processed_chunks
    gc.collect()
    
    print(f"[CHUNKED PROCESSING] Completed processing of '{sheet_name}': {len(final_result)} total rows")
    return final_result


def get_sheet_row_count(file_path, sheet_name):
    """
    Quickly get the row count of a sheet without loading all data.
    
    Args:
        file_path: Path to Excel file
        sheet_name: Name of sheet
    
    Returns:
        Number of rows in sheet (excluding header)
    """
    try:
        workbook = openpyxl.load_workbook(file_path, read_only=True)
        sheet = workbook[sheet_name]
        row_count = sheet.max_row - 1  # Subtract 1 for header row
        workbook.close()
        return row_count
    except Exception as e:
        print(f"[ROW COUNT] Error getting row count for '{sheet_name}': {e}")
        return 0


def process_uploaded_file(upload_session_id, file_path, original_filename, 
                          subscriber_id, subscriber_name, user_id,
                          reporting_month=None, reporting_year=None):
    """
    Asynchronous task to process uploaded Excel file
    
    Args:
        upload_session_id: ID of the UploadSession record
        file_path: Path to the uploaded file
        original_filename: Original name of uploaded file
        subscriber_id: Subscriber ID
        subscriber_name: Subscriber name
        user_id: User ID who uploaded the file
        reporting_month: User-selected reporting month (1-12) or None
        reporting_year: User-selected reporting year or None
    """
    print(f"[ASYNC TASK] Starting file processing for UploadSession {upload_session_id}")
    print(f"[ASYNC TASK] User-selected reporting period: {reporting_month}/{reporting_year}")
    
    try:
        # Get upload session
        upload_session = UploadSession.objects.get(id=upload_session_id)
        
        # Archive original file early in async processing (non-blocking)
        print(f"[ASYNC TASK] Archiving original file...")
        from .file_archival_utils import save_original_file
        try:
            with open(file_path, 'rb') as f:
                from django.core.files.base import File
                original_save_success, original_file_path, original_error = save_original_file(
                    File(f),
                    subscriber_id,
                    upload_session.uploaded_at
                )
                if original_save_success:
                    upload_session.original_file_path = original_file_path
                    upload_session.save(update_fields=['original_file_path'])
                    print(f"[ASYNC TASK] Original file archived: {original_file_path}")
                else:
                    print(f"[ASYNC TASK] Warning: Could not archive original file: {original_error}")
        except Exception as e:
            print(f"[ASYNC TASK] Warning: File archival error: {e}")
        
        # Update progress: Starting
        upload_session.update_progress('upload_validation', 5, 'File uploaded successfully, starting processing...')
        upload_session.mark_processing_started()
        
        # First, get sheet names and row counts without loading full data
        print(f"[ASYNC TASK] Analyzing Excel file structure: {file_path}")
        with pd.ExcelFile(file_path) as excel_file:
            sheet_names = excel_file.sheet_names
        
        print(f"[ASYNC TASK] Found {len(sheet_names)} sheets in workbook")
        
        # Get row counts for each sheet to determine processing strategy
        sheet_info = {}
        for sheet_name in sheet_names:
            row_count = get_sheet_row_count(file_path, sheet_name)
            sheet_info[sheet_name] = {
                'row_count': row_count,
                'use_chunking': row_count > LARGE_SHEET_THRESHOLD
            }
            if row_count > LARGE_SHEET_THRESHOLD:
                print(f"[ASYNC TASK] Large sheet detected: '{sheet_name}' ({row_count:,} rows) - will use chunked processing")
            else:
                print(f"[ASYNC TASK] Normal sheet: '{sheet_name}' ({row_count:,} rows)")
        
        upload_session.update_progress('upload_validation', 10, 'Excel file analyzed, loading sheets...')
        
        # Load sheets using appropriate method based on size
        xds = {}
        processing_stats = []
        
        for sheet_name in sheet_names:
            info = sheet_info[sheet_name]
            
            if info['use_chunking']:
                # For large sheets, we'll process them later in chunks
                # For now, just create a placeholder entry
                print(f"[ASYNC TASK] Skipping initial load for large sheet '{sheet_name}' - will process in chunks")
                xds[sheet_name] = None  # Placeholder - will be processed in chunks later
            else:
                # Load smaller sheets normally
                print(f"[ASYNC TASK] Loading sheet '{sheet_name}'")
                df = pd.read_excel(file_path, sheet_name=sheet_name, na_filter=False, dtype=object)
                
                # Convert all columns to string and clean nulls
                for col in df.columns:
                    df[col] = df[col].astype(str)
                    df[col] = df[col].replace({'nan': '', 'None': '', 'NaN': ''})
                
                xds[sheet_name] = df
            
            # Initialize processing stats
            processing_stats.append({
                'sheet_name': sheet_name,
                'initial_columns': 0,  # Will be updated during processing
                'initial_records': info['row_count'],
                'processed_columns': None,
                'valid_records': 0
            })
        
        # Ensure all required sheets exist
        processed_sheets = ensure_all_sheets_exist(xds)
        
        upload_session.update_progress('data_mapping', 15, 'Sheets validated, starting processing...')
        
        # Process each sheet using appropriate method (chunked vs normal)
        total_sheets = len(sheet_names)
        for idx, sheet_name in enumerate(sheet_names):
            progress = 15 + (idx * 40 // total_sheets)  # Progress from 15% to 55%
            cleaned_name = clean_sheet_name(sheet_name)
            info = sheet_info[sheet_name]
            
            if info['use_chunking']:
                # Process large sheet in chunks
                upload_session.update_progress(
                    'data_mapping' if progress < 30 else 'data_cleaning',
                    progress,
                    f'Processing large sheet {sheet_name} ({info["row_count"]:,} rows) in chunks...'
                )
                
                processed_df = process_large_sheet_chunked(
                    file_path, 
                    sheet_name, 
                    cleaned_name, 
                    chunk_size=CHUNK_SIZE
                )
                
                processed_sheets[cleaned_name] = processed_df
                
                # Update stats
                for stat in processing_stats:
                    if stat['sheet_name'] == sheet_name:
                        stat['processed_columns'] = len(processed_df.columns)
                        stat['valid_records'] = len(processed_df)
                        break
                
                # Force garbage collection after large sheet
                gc.collect()
                
            else:
                # Process normal-sized sheet
                upload_session.update_progress(
                    'data_mapping' if progress < 30 else 'data_cleaning',
                    progress,
                    f'Processing sheet {sheet_name}...'
                )
                
                sheet_data = xds[sheet_name]
                if sheet_data is not None:
                    processed_df = process_single_sheet(sheet_data, cleaned_name)
                    processed_sheets[cleaned_name] = processed_df
                    
                    # Update stats
                    for stat in processing_stats:
                        if stat['sheet_name'] == sheet_name:
                            stat['initial_columns'] = len(sheet_data.columns)
                            stat['processed_columns'] = len(processed_df.columns)
                            stat['valid_records'] = len(processed_df)
                            break
        
        # ========================================================================
        # VALIDATE REQUIRED COLUMNS (After Renaming, Before Processing)
        # ========================================================================
        upload_session.update_progress('data_cleaning', 54, 'Validating required columns...')
        
        print(f"[VALIDATION] Starting column validation for {len(processed_sheets)} sheets")
        is_valid, validation_message = validate_required_columns(processed_sheets)
        
        if not is_valid:
            # Validation failed - mark session as failed and stop processing
            print(f"[VALIDATION] ❌ Validation failed: {validation_message}")
            
            upload_session.update_progress(
                'upload_validation', 
                0, 
                'Validation Failed: Required columns are empty'
            )
            
            # Log detailed error to activity log
           
            upload_session.activity_log = json.dumps({
                'validation_error': validation_message,
                'timestamp': dt.now().isoformat()
            })
            
            upload_session.mark_failed(validation_message)
            upload_session.save()
            
            # Clean up the uploaded file
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    print(f"[CLEANUP] Removed uploaded file: {file_path}")
                except Exception as e:
                    print(f"[CLEANUP] Failed to remove file: {e}")
            
            return  # Stop processing
        
        print(f"[VALIDATION] ✓ {validation_message}")
        upload_session.update_progress('credit_matching', 55, 'Validation successful, saving to temporary storage...')
        
        # ========================================================================
        # SAVE PROCESSED SHEETS TO PARQUET (Clear memory)
        # ========================================================================
        
        # Track which sheets were actually saved to Parquet
        saved_to_parquet = []
        
        # Save each processed sheet to Parquet and clear from memory
        # IMPORTANT: Also save empty auto-generated sheets to preserve headers for merging
        for sheet_key, sheet_df in processed_sheets.items():
            if sheet_df is not None:
                # Save both non-empty processed sheets AND empty auto-generated sheets (with headers)
                # Empty auto-generated sheets are crucial for merge operations to have proper columns
                save_to_parquet(sheet_df, upload_session_id, sheet_key)
                saved_to_parquet.append(sheet_key)
                processed_sheets[sheet_key] = None  # Clear from memory
                
        # Force garbage collection after clearing all sheets
        gc.collect()
        print(f"[MEMORY] Cleared all processed sheets from memory, saved to Parquet")
        print(f"[MEMORY] Sheets saved to Parquet (including auto-generated with headers): {saved_to_parquet}")
        
        # Initialize credit DataFrame (will be empty for merged sheets)
        credit = pd.DataFrame()
        
        upload_session.update_progress('credit_matching', 58, 'Matching borrowers with credits...')
        
        # Handle merged sheets or perform merging
        if 'consumermerged' in saved_to_parquet or 'commercialmerged' in saved_to_parquet:
            # Load merged sheets from Parquet (only when needed)
            indi = load_from_parquet(upload_session_id, 'consumermerged') if 'consumermerged' in saved_to_parquet else pd.DataFrame()
            corpo = load_from_parquet(upload_session_id, 'commercialmerged') if 'commercialmerged' in saved_to_parquet else pd.DataFrame()
            individual_credit_unmerged = {'unmerged_count': 0, 'total_valid_credit': 0, 'matched_count': 0}
            corporate_credit_unmerged = {'unmerged_count': 0, 'total_valid_credit': 0, 'matched_count': 0}
        else:
            # Load only the sheets that were actually saved to Parquet
            print("[MEMORY] Loading sheets from Parquet for merging...")
            
            consu = load_from_parquet(upload_session_id, 'individualborrowertemplate') if 'individualborrowertemplate' in saved_to_parquet else pd.DataFrame()
            comm = load_from_parquet(upload_session_id, 'corporateborrowertemplate') if 'corporateborrowertemplate' in saved_to_parquet else pd.DataFrame()
            credit = load_from_parquet(upload_session_id, 'creditinformation') if 'creditinformation' in saved_to_parquet else pd.DataFrame()
            guar = load_from_parquet(upload_session_id, 'guarantorsinformation') if 'guarantorsinformation' in saved_to_parquet else pd.DataFrame()
            prin = load_from_parquet(upload_session_id, 'principalofficerstemplate') if 'principalofficerstemplate' in saved_to_parquet else pd.DataFrame()
            
            individual_credit_unmerged = {'unmerged_count': 0, 'total_valid_credit': 0, 'matched_count': 0}
            corporate_credit_unmerged = {'unmerged_count': 0, 'total_valid_credit': 0, 'matched_count': 0}
            
            # Perform merging and track matched credit IDs
            upload_session.update_progress('credit_matching', 65, 'Merging individual borrowers with credits...')
            indi, individual_matched_credit_ids = merge_individual_borrowers(consu, credit, guar)
            
            # Clear merged source data from memory (keep credit for now to identify unmatched)
            del consu, guar
            gc.collect()
            
            upload_session.update_progress('credit_matching', 70, 'Merging corporate borrowers with credits...')
            corpo, corporate_matched_credit_ids = merge_corporate_borrowers(comm, credit, prin)
            
            # Clear merged source data from memory (keep credit for now to identify unmatched)
            del comm, prin
            gc.collect()
            
            # Identify truly unmatched credits (didn't match either individual or corporate)
            upload_session.update_progress('credit_matching', 73, 'Identifying unmatched credits...')
            
            # Combine all matched credit IDs from both merges
            all_matched_credit_ids = individual_matched_credit_ids | corporate_matched_credit_ids
            
            # Find credits that were not matched in either merge
            actual_unmatched_count = 0  # Initialize the ground truth count
            if not credit.empty and 'CUSTOMERID' in credit.columns:
                unmatched_credits = credit[~credit['CUSTOMERID'].isin(all_matched_credit_ids)].copy()
                
                # Keep only the 5 required columns for the report
                required_columns = ['CUSTOMERID', 'ACCOUNTNUMBER', 'ACCOUNTSTATUS', 'CREDITLIMIT', 'AVAILEDLIMIT']
                available_columns = [col for col in required_columns if col in unmatched_credits.columns]
                
                unmatched_credits_report = unmatched_credits[available_columns].copy()
                
                # Store the actual count (ground truth)
                actual_unmatched_count = len(unmatched_credits_report)
                
                print(f"[UNMATCHED CREDITS] Total credits: {len(credit)}")
                print(f"[UNMATCHED CREDITS] Matched to individual: {len(individual_matched_credit_ids)}")
                print(f"[UNMATCHED CREDITS] Matched to corporate: {len(corporate_matched_credit_ids)}")
                print(f"[UNMATCHED CREDITS] Truly unmatched (GROUND TRUTH): {actual_unmatched_count}")
                
                # Save unmatched credits to Parquet (will be preserved for Excel report)
                if not unmatched_credits_report.empty:
                    save_to_parquet(unmatched_credits_report, upload_session_id, 'unmatched_credits')
                    print(f"[UNMATCHED CREDITS] Saved {actual_unmatched_count} unmatched credits to Parquet")
                else:
                    # Save empty DataFrame with correct columns to Parquet
                    empty_unmatched = pd.DataFrame(columns=available_columns)
                    save_to_parquet(empty_unmatched, upload_session_id, 'unmatched_credits')
                    print(f"[UNMATCHED CREDITS] No unmatched credits - saved empty DataFrame with headers")
            else:
                # No credit data available
                empty_unmatched = pd.DataFrame(columns=['CUSTOMERID', 'ACCOUNTNUMBER', 'ACCOUNTSTATUS', 'CREDITLIMIT', 'AVAILEDLIMIT'])
                save_to_parquet(empty_unmatched, upload_session_id, 'unmatched_credits')
                print(f"[UNMATCHED CREDITS] No credit data - saved empty DataFrame")
            
            # Store the actual unmatched count in upload session for result page display
            upload_session.unmatched_credit_records = actual_unmatched_count
            upload_session.save()
            print(f"[UNMATCHED CREDITS] Stored ground truth count in upload_session: {actual_unmatched_count}")
            
            # Now safe to delete credit DataFrame
            del credit
            gc.collect()
        
        upload_session.update_progress('credit_matching', 75, 'Credit matching complete, identifying split candidates...')
        
        # Save merged data to Parquet before splitting (needed for post-verification)
        save_to_parquet(indi, upload_session_id, 'merged_individual')
        save_to_parquet(corpo, upload_session_id, 'merged_corporate')
        
        # Split commercial/consumer entities for human verification
        split_indi, split_candidates_commercial = split_commercial_entities(indi)
        split_corpo, split_candidates_consumer = split_consumer_entities(corpo)
        
        # Clear original merged data from memory (keep only split versions)
        del indi, corpo
        gc.collect()
        
        # Save split data to Parquet (needed for post-verification final output)
        save_to_parquet(split_indi, upload_session_id, 'split_individual')
        save_to_parquet(split_corpo, upload_session_id, 'split_corporate')
        
        # Reorder columns for display
        reordered_commercial_columns = reorder_commercial_columns(split_candidates_commercial.columns)
        split_candidates_commercial = split_candidates_commercial.reindex(columns=reordered_commercial_columns)
        
        reordered_consumer_columns = reorder_consumer_columns(split_candidates_consumer.columns)
        split_candidates_consumer = split_candidates_consumer.reindex(columns=reordered_consumer_columns)
        
        # Prepare verification data (lightweight - only candidates for UI)
        commercial_records = json.loads(split_candidates_commercial.to_json(orient='records'))
        consumer_records = json.loads(split_candidates_consumer.to_json(orient='records'))
        
        upload_session.update_progress('awaiting_verification', 80, 'Processing complete, awaiting human verification...')
        
        # Determine final reporting period with priority system
        # Priority: 1) User-selected, 2) Filename extraction, 3) Previous month default
        from datetime import datetime
        extracted_month, extracted_year = extract_date_from_filename(original_filename)
        
        now = datetime.now()
        default_month = now.month - 1 if now.month > 1 else 12
        default_year = now.year if now.month > 1 else now.year - 1
        
        final_month = reporting_month or extracted_month or default_month
        final_year = reporting_year or extracted_year or default_year
        
        print(f"[DATE DETERMINATION] User-selected: {reporting_month}/{reporting_year}")
        print(f"[DATE DETERMINATION] Filename-extracted: {extracted_month}/{extracted_year}")
        print(f"[DATE DETERMINATION] Final decision: {final_month}/{final_year}")
        
        # Store ONLY verification candidates in JSON (not full DataFrames)
        # Full processed data is in Parquet files
        upload_session.processing_data = json.dumps({
            # Small data only - candidates for verification UI
            'commercial_candidates': commercial_records,
            'consumer_candidates': consumer_records,
            'columns_commercial': list(reordered_commercial_columns),
            'columns_consumer': list(reordered_consumer_columns),
            
            # Metadata
            'processing_stats': json.loads(json.dumps(processing_stats, default=convert_numpy)),
            'original_filename': original_filename,
            'subscriber_id': subscriber_id,
            'subscriber_name': subscriber_name,
            
            # Reporting period (final decision)
            'final_month': final_month,
            'final_year': final_year,
            
            # Credit stats (small)
            'individual_credit_unmerged': {
                k: int(v) if isinstance(v, (np.int64, np.int32)) else v 
                for k, v in individual_credit_unmerged.items()
            },
            'corporate_credit_unmerged': {
                k: int(v) if isinstance(v, (np.int64, np.int32)) else v 
                for k, v in corporate_credit_unmerged.items()
            },
            
            # Flag indicating Parquet files are available
            'using_parquet': True
        })
        
        # Mark as awaiting verification
        upload_session.mark_awaiting_verification(
            commercial_count=len(commercial_records),
            consumer_count=len(consumer_records)
        )
        
        upload_session.save()
        
        print(f"[ASYNC TASK] Processing complete for UploadSession {upload_session_id}")
        print(f"[ASYNC TASK] Awaiting human verification: {len(commercial_records)} commercial, {len(consumer_records)} consumer candidates")
        
    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        print(f"[ASYNC TASK ERROR] {error_details}")
        
        # Mark upload session as failed
        try:
            upload_session = UploadSession.objects.get(id=upload_session_id)
            upload_session.mark_failed(error_message=str(e))
            
            # Cleanup temp Parquet files on error
            cleanup_temp_files(upload_session_id)
        except:
            pass
        
        raise
    
    finally:
        # Clean up temp uploaded file
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
                print(f"[ASYNC TASK] Removed temp file: {file_path}")
            except Exception as e:
                print(f"[ASYNC TASK] Failed to remove temp file: {e}")


def process_post_verification(upload_session_id, use_parquet=True,
                               processing_stats=None, original_filename=None,
                               subscriber_id=None, subscriber_name=None,
                               total_individual_records=0, total_corporate_records=0):
    """
    Asynchronous task to process post-verification file generation
    
    Args:
        upload_session_id: ID of the UploadSession record
        use_parquet: Whether to load data from Parquet files (True) or JSON (False - legacy)
        processing_stats: Processing statistics
        original_filename: Original filename for naming outputs
        subscriber_id: Subscriber ID
        subscriber_name: Subscriber name
        total_individual_records: Count of individual records
        total_corporate_records: Count of corporate records
    """
    from .views import (
        modify_middle_names, clean_for_output, remove_duplicates,
        generate_filename_from_user, extract_date_from_filename,
        save_excel_as_text, enforce_string_columns
    )
    
    try:
        print(f"[POST-VERIFICATION TASK] Starting post-verification processing for UploadSession {upload_session_id}")
        
        upload_session = UploadSession.objects.get(id=upload_session_id)
        
        # Update progress to post-verification stage
        upload_session.update_progress('post_verification', 85, 'Processing verification results...')
        
        # Load DataFrames from Parquet (memory efficient) or JSON (legacy)
        if use_parquet:
            print("[POST-VERIFICATION] Loading final data from Parquet files")
            indi = load_from_parquet(upload_session_id, 'split_individual')
            corpo = load_from_parquet(upload_session_id, 'split_corporate')
            indi = enforce_string_columns(indi)
            corpo = enforce_string_columns(corpo)
        else:
            # Legacy: Load from JSON (for backward compatibility)
            print("[POST-VERIFICATION] Loading final data from JSON (legacy mode)")
            # This would be passed in from the calling function
            # Not implemented fully as we're moving to Parquet
            raise NotImplementedError("JSON loading not implemented - use Parquet")
        
        # Apply final transformations
        upload_session.update_progress('post_verification', 88, 'Applying final data transformations...')
        indi = modify_middle_names(indi)
        corpo = modify_middle_names(corpo)
        
        indi = clean_for_output(indi)
        corpo = clean_for_output(corpo)
        
        # Drop name and dependant columns from corpo
        columns_to_remove = ['SURNAME', 'FIRSTNAME', 'MIDDLENAME', 'DEPENDANTS']
        corpo = corpo.drop(columns=[col for col in columns_to_remove if col in corpo.columns], errors='ignore')
        
        indi = remove_duplicates(indi)
        corpo = remove_duplicates(corpo)
        
        # Note: unmatched_credit_records was already calculated and stored during process_uploaded_file()
        # We preserve that ground truth value instead of recalculating here
        upload_session.update_progress('output_generation', 90, 'Generating output files...')
        
        # ========================================================================
        # GET REPORTING PERIOD (from stored processing_data)
        # ========================================================================
        # Retrieve the final month/year that was determined during process_uploaded_file
        # This uses the priority system: user-selected > filename > default
        upload_session_obj = UploadSession.objects.get(id=upload_session_id)
        processing_data_json = json.loads(upload_session_obj.processing_data)
        
        extracted_month = processing_data_json.get('final_month')
        extracted_year = processing_data_json.get('final_year')
        
        print(f"[OUTPUT] Using reporting period from processing_data: {extracted_month}/{extracted_year}")
        
        # ========================================================================
        # SMART OUTPUT FORMAT SELECTION
        # ========================================================================
        # Configuration
        EXCEL_ROW_THRESHOLD = 300000  # Switch to TXT if sheet exceeds this
        
        # Generate output files with smart format selection
        upload_session.update_progress('output_generation', 92, 'Generating outputs with smart format selection...')
        
        # Create subscriber-specific output directory
        subscriber_name_clean = subscriber_name.replace(' ', '')
        subscriber_output_dir = os.path.join(settings.MEDIA_ROOT, 'outputs', subscriber_name_clean)
        os.makedirs(subscriber_output_dir, exist_ok=True)
        
        fs = FileSystemStorage()
        
        # ========================================================================
        # INDIVIDUAL SHEET - Smart Format Selection
        # ========================================================================
        individual_row_count = len(indi)
        
        if individual_row_count > EXCEL_ROW_THRESHOLD:
            # Generate TXT for large individual sheet
            print(f"[OUTPUT] Individual sheet has {individual_row_count} rows - generating TXT (threshold: {EXCEL_ROW_THRESHOLD})")
            indi_output_filename = generate_filename_from_user(subscriber_id, subscriber_name, 'txt', 'consumer', extracted_month, extracted_year)
            indi_output_path = os.path.join(subscriber_output_dir, indi_output_filename)
            indi.to_csv(indi_output_path, sep='\t', index=False, encoding='utf-8')
            indi_output_format = 'txt'
            print(f"[OUTPUT] Individual TXT generated: {indi_output_filename}")
        else:
            # Generate Excel for small individual sheet
            print(f"[OUTPUT] Individual sheet has {individual_row_count} rows - generating Excel (threshold: {EXCEL_ROW_THRESHOLD})")
            indi_output_filename = generate_filename_from_user(subscriber_id, subscriber_name, 'excel', 'consumer', extracted_month, extracted_year)
            indi_output_path = os.path.join(subscriber_output_dir, indi_output_filename)
            save_excel_as_text(indi, indi_output_path)
            indi_output_format = 'xlsx'
            print(f"[OUTPUT] Individual Excel generated: {indi_output_filename}")
        
        # Store relative path for URL generation
        indi_output_relative = os.path.join('outputs', subscriber_name_clean, indi_output_filename)
        indi_processed_file_url = fs.url(indi_output_relative)
        
        # ========================================================================
        # CORPORATE SHEET - Smart Format Selection
        # ========================================================================
        corporate_row_count = len(corpo)
        
        if corporate_row_count > EXCEL_ROW_THRESHOLD:
            # Generate TXT for large corporate sheet
            print(f"[OUTPUT] Corporate sheet has {corporate_row_count} rows - generating TXT (threshold: {EXCEL_ROW_THRESHOLD})")
            corpo_output_filename = generate_filename_from_user(subscriber_id, subscriber_name, 'txt', 'commercial', extracted_month, extracted_year)
            corpo_output_path = os.path.join(subscriber_output_dir, corpo_output_filename)
            corpo.to_csv(corpo_output_path, sep='\t', index=False, encoding='utf-8')
            corpo_output_format = 'txt'
            print(f"[OUTPUT] Corporate TXT generated: {corpo_output_filename}")
        else:
            # Generate Excel for small corporate sheet
            print(f"[OUTPUT] Corporate sheet has {corporate_row_count} rows - generating Excel (threshold: {EXCEL_ROW_THRESHOLD})")
            corpo_output_filename = generate_filename_from_user(subscriber_id, subscriber_name, 'excel', 'commercial', extracted_month, extracted_year)
            corpo_output_path = os.path.join(subscriber_output_dir, corpo_output_filename)
            save_excel_as_text(corpo, corpo_output_path)
            corpo_output_format = 'xlsx'
            print(f"[OUTPUT] Corporate Excel generated: {corpo_output_filename}")
        
        # Store relative path for URL generation
        corpo_output_relative = os.path.join('outputs', subscriber_name_clean, corpo_output_filename)
        corpo_processed_file_url = fs.url(corpo_output_relative)
        
        # ========================================================================
        # UPDATE UPLOAD SESSION
        # ========================================================================
        upload_session.update_progress('output_generation', 98, 'Finalizing upload session...')
        
        upload_session.individual_file_path = indi_processed_file_url
        upload_session.corporate_file_path = corpo_processed_file_url
        # Store output formats for download UI
        upload_session.individual_output_format = indi_output_format
        upload_session.corporate_output_format = corpo_output_format
        # Note: unmatched_credit_records already set during process_uploaded_file() - don't overwrite
        upload_session.individual_credit_matched = total_individual_records
        upload_session.corporate_credit_matched = total_corporate_records
        
        # Mark as completed
        upload_session.mark_completed(
            individual_count=total_individual_records,
            corporate_count=total_corporate_records,
            processing_time=None  # Can calculate if needed
        )
        
        upload_session.save()
        
        # ========================================================================
        # CLEANUP: Delete temporary Parquet files after successful completion
        # ========================================================================
        cleanup_temp_files(upload_session_id)
        
        print(f"[POST-VERIFICATION TASK] Processing complete for UploadSession {upload_session_id}")
        print(f"[POST-VERIFICATION TASK] Files generated: {indi_output_filename}, {corpo_output_filename}")
        
    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        print(f"[POST-VERIFICATION TASK ERROR] {error_details}")
        
        # Mark upload session as failed
        try:
            upload_session = UploadSession.objects.get(id=upload_session_id)
            upload_session.mark_failed(error_message=str(e))
            
            # Cleanup temp files even on error
            cleanup_temp_files(upload_session_id)
        except:
            pass
        
        raise

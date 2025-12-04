"""
Utility functions for original file archival system.

This module provides functions to store and retrieve original uploaded files
for dispute resolution, compliance, and audit purposes.
"""

import os
import shutil
from datetime import datetime
from django.conf import settings
from django.core.files.storage import default_storage
from django.utils import timezone
import logging

logger = logging.getLogger(__name__)


def generate_original_filename(subscriber_id, original_filename, upload_timestamp=None):
    """
    Generate a unique filename for storing original files.
    
    Format: {subscriber_id}_{timestamp}_{original_filename}
    
    Args:
        subscriber_id (int): The subscriber ID
        original_filename (str): The original filename from upload
        upload_timestamp (datetime, optional): Upload timestamp. Defaults to current time.
    
    Returns:
        str: Generated unique filename
    """
    if upload_timestamp is None:
        upload_timestamp = timezone.now()
    
    # Format timestamp as YYYYMMDD_HHMMSS
    timestamp_str = upload_timestamp.strftime("%Y%m%d_%H%M%S")
    
    # Clean the original filename to remove any path separators
    clean_filename = os.path.basename(original_filename)
    
    # Generate the new filename
    new_filename = f"{subscriber_id}_{timestamp_str}_{clean_filename}"
    
    return new_filename


def get_original_file_path(filename):
    """
    Get the full path for storing an original file.
    
    Args:
        filename (str): The filename to store
    
    Returns:
        str: Full path where the file should be stored
    """
    originals_dir = os.path.join(settings.MEDIA_ROOT, 'originals')
    return os.path.join(originals_dir, filename)


def save_original_file(uploaded_file, subscriber_id, upload_timestamp=None):
    """
    Save the original uploaded file to the originals directory.
    
    Args:
        uploaded_file: Django UploadedFile object
        subscriber_id (int): The subscriber ID
        upload_timestamp (datetime, optional): Upload timestamp. Defaults to current time.
    
    Returns:
        tuple: (success: bool, file_path: str, error_message: str)
    """
    try:
        # Generate unique filename
        original_filename = generate_original_filename(
            subscriber_id, 
            uploaded_file.name, 
            upload_timestamp
        )
        
        # Get the full path
        file_path = get_original_file_path(original_filename)
        
        # Ensure the originals directory exists
        originals_dir = os.path.dirname(file_path)
        os.makedirs(originals_dir, exist_ok=True)
        
        # Save the file
        with open(file_path, 'wb') as destination:
            for chunk in uploaded_file.chunks():
                destination.write(chunk)
        
        logger.info(f"Original file saved successfully: {file_path}")
        return True, file_path, None
        
    except Exception as e:
        error_msg = f"Failed to save original file: {str(e)}"
        logger.error(error_msg)
        return False, None, error_msg


def copy_file_to_originals(source_file_path, subscriber_id, original_filename, upload_timestamp=None):
    """
    Copy an existing file to the originals directory.
    
    This is useful when you already have a file saved elsewhere and want to archive it.
    
    Args:
        source_file_path (str): Path to the source file
        subscriber_id (int): The subscriber ID
        original_filename (str): The original filename
        upload_timestamp (datetime, optional): Upload timestamp. Defaults to current time.
    
    Returns:
        tuple: (success: bool, file_path: str, error_message: str)
    """
    try:
        if not os.path.exists(source_file_path):
            return False, None, f"Source file does not exist: {source_file_path}"
        
        # Generate unique filename
        archive_filename = generate_original_filename(
            subscriber_id, 
            original_filename, 
            upload_timestamp
        )
        
        # Get the full path
        destination_path = get_original_file_path(archive_filename)
        
        # Ensure the originals directory exists
        originals_dir = os.path.dirname(destination_path)
        os.makedirs(originals_dir, exist_ok=True)
        
        # Copy the file
        shutil.copy2(source_file_path, destination_path)
        
        logger.info(f"File copied to originals: {destination_path}")
        return True, destination_path, None
        
    except Exception as e:
        error_msg = f"Failed to copy file to originals: {str(e)}"
        logger.error(error_msg)
        return False, None, error_msg


def retrieve_original_file(file_path):
    """
    Retrieve an original file for download or verification.
    
    Args:
        file_path (str): Path to the original file
    
    Returns:
        tuple: (success: bool, file_exists: bool, full_path: str, error_message: str)
    """
    try:
        # If it's a relative path, make it absolute
        if not os.path.isabs(file_path):
            full_path = os.path.join(settings.MEDIA_ROOT, file_path)
        else:
            full_path = file_path
        
        # Check if file exists
        file_exists = os.path.exists(full_path)
        
        if file_exists:
            logger.info(f"Original file found: {full_path}")
            return True, True, full_path, None
        else:
            logger.warning(f"Original file not found: {full_path}")
            return True, False, full_path, "File not found"
        
    except Exception as e:
        error_msg = f"Error retrieving original file: {str(e)}"
        logger.error(error_msg)
        return False, False, None, error_msg


def delete_original_file(file_path):
    """
    Delete an original file (use with caution - for cleanup purposes only).
    
    Args:
        file_path (str): Path to the original file
    
    Returns:
        tuple: (success: bool, error_message: str)
    """
    try:
        # If it's a relative path, make it absolute
        if not os.path.isabs(file_path):
            full_path = os.path.join(settings.MEDIA_ROOT, file_path)
        else:
            full_path = file_path
        
        if os.path.exists(full_path):
            os.remove(full_path)
            logger.info(f"Original file deleted: {full_path}")
            return True, None
        else:
            logger.warning(f"Original file not found for deletion: {full_path}")
            return True, "File not found"
        
    except Exception as e:
        error_msg = f"Failed to delete original file: {str(e)}"
        logger.error(error_msg)
        return False, error_msg


def get_original_file_info(file_path):
    """
    Get information about an original file.
    
    Args:
        file_path (str): Path to the original file
    
    Returns:
        dict: File information including size, creation time, etc.
    """
    try:
        # If it's a relative path, make it absolute
        if not os.path.isabs(file_path):
            full_path = os.path.join(settings.MEDIA_ROOT, file_path)
        else:
            full_path = file_path
        
        if not os.path.exists(full_path):
            return {'exists': False, 'error': 'File not found'}
        
        stat = os.stat(full_path)
        
        return {
            'exists': True,
            'size': stat.st_size,
            'created': datetime.fromtimestamp(stat.st_ctime),
            'modified': datetime.fromtimestamp(stat.st_mtime),
            'full_path': full_path,
            'filename': os.path.basename(full_path)
        }
        
    except Exception as e:
        return {'exists': False, 'error': str(e)}


def cleanup_old_originals(days_to_keep=365):
    """
    Clean up original files older than specified days.
    
    This function should be used carefully and typically run as a scheduled task.
    Consider compliance requirements before implementing automatic cleanup.
    
    Args:
        days_to_keep (int): Number of days to keep files. Default is 365 (1 year).
    
    Returns:
        dict: Cleanup statistics
    """
    try:
        originals_dir = os.path.join(settings.MEDIA_ROOT, 'originals')
        
        if not os.path.exists(originals_dir):
            return {'error': 'Originals directory does not exist'}
        
        cutoff_time = timezone.now().timestamp() - (days_to_keep * 24 * 60 * 60)
        
        deleted_count = 0
        total_size_freed = 0
        errors = []
        
        for filename in os.listdir(originals_dir):
            file_path = os.path.join(originals_dir, filename)
            
            if os.path.isfile(file_path):
                try:
                    file_stat = os.stat(file_path)
                    if file_stat.st_mtime < cutoff_time:
                        file_size = file_stat.st_size
                        os.remove(file_path)
                        deleted_count += 1
                        total_size_freed += file_size
                        logger.info(f"Deleted old original file: {file_path}")
                except Exception as e:
                    errors.append(f"Failed to delete {filename}: {str(e)}")
        
        return {
            'deleted_count': deleted_count,
            'total_size_freed': total_size_freed,
            'errors': errors,
            'success': True
        }
        
    except Exception as e:
        logger.error(f"Error during cleanup: {str(e)}")
        return {'error': str(e), 'success': False}
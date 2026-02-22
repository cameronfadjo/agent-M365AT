"""
Blob Storage Module - Refresh v5
Handles document storage and SAS URL generation for Power Automate compatibility.

This module solves the binary corruption issue that occurs when passing
Word documents through Power Automate to SharePoint. Instead of returning
base64-encoded content, documents are uploaded to Azure Blob Storage and
a secure SAS URL is returned for direct download.

Environment Variables Required:
  AZURE_STORAGE_CONNECTION_STRING  - Full connection string (preferred)
  OR
  AZURE_STORAGE_ACCOUNT_NAME       - Storage account name
  AZURE_STORAGE_ACCOUNT_KEY        - Storage account key

  AZURE_STORAGE_CONTAINER_NAME     - Container name (default: "generated-documents")
  SAS_TOKEN_EXPIRY_HOURS           - SAS URL validity period (default: 24)
"""

import os
import logging
from datetime import datetime, timedelta, timezone
from typing import Optional, Tuple
import uuid

# Azure Storage imports - will be available after deployment
try:
    from azure.storage.blob import (
        BlobServiceClient,
        BlobClient,
        generate_blob_sas,
        BlobSasPermissions,
        ContentSettings
    )
    BLOB_STORAGE_AVAILABLE = True
except ImportError:
    BLOB_STORAGE_AVAILABLE = False
    logging.warning("azure-storage-blob not installed. SAS URL functionality disabled.")


def is_blob_storage_configured() -> bool:
    """Check if blob storage is properly configured."""
    if not BLOB_STORAGE_AVAILABLE:
        return False

    # Check for connection string
    if os.environ.get('AZURE_STORAGE_CONNECTION_STRING'):
        return True

    # Check for account name + key
    if (os.environ.get('AZURE_STORAGE_ACCOUNT_NAME') and
        os.environ.get('AZURE_STORAGE_ACCOUNT_KEY')):
        return True

    return False


def get_blob_service_client() -> Optional['BlobServiceClient']:
    """Get Azure Blob Service client from environment configuration."""
    if not BLOB_STORAGE_AVAILABLE:
        return None

    # Try connection string first
    connection_string = os.environ.get('AZURE_STORAGE_CONNECTION_STRING')
    if connection_string:
        return BlobServiceClient.from_connection_string(connection_string)

    # Fall back to account name + key
    account_name = os.environ.get('AZURE_STORAGE_ACCOUNT_NAME')
    account_key = os.environ.get('AZURE_STORAGE_ACCOUNT_KEY')

    if account_name and account_key:
        account_url = f"https://{account_name}.blob.core.windows.net"
        return BlobServiceClient(account_url=account_url, credential=account_key)

    return None


def get_container_name() -> str:
    """Get the container name for document storage."""
    return os.environ.get('AZURE_STORAGE_CONTAINER_NAME', 'generated-documents')


def get_sas_expiry_hours() -> int:
    """Get the SAS token expiry time in hours."""
    try:
        return int(os.environ.get('SAS_TOKEN_EXPIRY_HOURS', '24'))
    except ValueError:
        return 24


def upload_document_and_get_sas_url(
    document_bytes: bytes,
    filename: str,
    content_type: str = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
) -> Tuple[bool, str, Optional[str]]:
    """
    Upload a document to Azure Blob Storage and return a SAS URL.

    Args:
        document_bytes: The document content as bytes
        filename: The desired filename for download
        content_type: MIME type of the document

    Returns:
        Tuple of (success: bool, message_or_url: str, blob_name: Optional[str])
        - On success: (True, sas_url, blob_name)
        - On failure: (False, error_message, None)
    """
    if not BLOB_STORAGE_AVAILABLE:
        return (False, "Azure Blob Storage SDK not available. Install azure-storage-blob.", None)

    if not is_blob_storage_configured():
        return (False, "Azure Blob Storage not configured. Set AZURE_STORAGE_CONNECTION_STRING or AZURE_STORAGE_ACCOUNT_NAME + AZURE_STORAGE_ACCOUNT_KEY.", None)

    try:
        blob_service_client = get_blob_service_client()
        if not blob_service_client:
            return (False, "Failed to create Blob Service client.", None)

        container_name = get_container_name()

        # Ensure container exists
        container_client = blob_service_client.get_container_client(container_name)
        try:
            container_client.get_container_properties()
        except Exception:
            # Container doesn't exist, create it
            container_client.create_container()
            logging.info(f"Created container: {container_name}")

        # Generate unique blob name with timestamp and UUID
        timestamp = datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')
        unique_id = str(uuid.uuid4())[:8]
        # Sanitize filename
        safe_filename = "".join(c for c in filename if c.isalnum() or c in '._- ')
        blob_name = f"{timestamp}_{unique_id}_{safe_filename}"

        # Upload the document
        blob_client = container_client.get_blob_client(blob_name)

        content_settings = ContentSettings(
            content_type=content_type,
            content_disposition=f'attachment; filename="{filename}"'
        )

        blob_client.upload_blob(
            document_bytes,
            overwrite=True,
            content_settings=content_settings
        )

        logging.info(f"Uploaded document to blob: {blob_name}")

        # Generate SAS URL
        sas_url = generate_sas_url(blob_name, filename)

        if sas_url:
            return (True, sas_url, blob_name)
        else:
            # Fallback: return blob URL without SAS (requires public access)
            blob_url = blob_client.url
            return (True, blob_url, blob_name)

    except Exception as e:
        logging.error(f"Error uploading to blob storage: {str(e)}")
        return (False, f"Blob storage error: {str(e)}", None)


def generate_sas_url(blob_name: str, download_filename: str = None) -> Optional[str]:
    """
    Generate a SAS URL for a blob with read permission.

    Args:
        blob_name: The name of the blob in storage
        download_filename: Optional filename for Content-Disposition

    Returns:
        The SAS URL string, or None if generation fails
    """
    if not BLOB_STORAGE_AVAILABLE:
        return None

    try:
        account_name = os.environ.get('AZURE_STORAGE_ACCOUNT_NAME')
        account_key = os.environ.get('AZURE_STORAGE_ACCOUNT_KEY')

        # If using connection string, parse account name and key
        connection_string = os.environ.get('AZURE_STORAGE_CONNECTION_STRING')
        if connection_string and not (account_name and account_key):
            # Parse connection string
            parts = dict(part.split('=', 1) for part in connection_string.split(';') if '=' in part)
            account_name = parts.get('AccountName')
            account_key = parts.get('AccountKey')

        if not account_name or not account_key:
            logging.warning("Cannot generate SAS URL: missing account credentials")
            return None

        container_name = get_container_name()
        expiry_hours = get_sas_expiry_hours()

        # Generate SAS token
        sas_token = generate_blob_sas(
            account_name=account_name,
            container_name=container_name,
            blob_name=blob_name,
            account_key=account_key,
            permission=BlobSasPermissions(read=True),
            expiry=datetime.now(timezone.utc) + timedelta(hours=expiry_hours),
            content_disposition=f'attachment; filename="{download_filename}"' if download_filename else None
        )

        # Construct full URL
        sas_url = f"https://{account_name}.blob.core.windows.net/{container_name}/{blob_name}?{sas_token}"

        logging.info(f"Generated SAS URL for blob: {blob_name}, expires in {expiry_hours} hours")
        return sas_url

    except Exception as e:
        logging.error(f"Error generating SAS URL: {str(e)}")
        return None


def delete_blob(blob_name: str) -> bool:
    """
    Delete a blob from storage (cleanup utility).

    Args:
        blob_name: The name of the blob to delete

    Returns:
        True if deleted successfully, False otherwise
    """
    if not BLOB_STORAGE_AVAILABLE or not is_blob_storage_configured():
        return False

    try:
        blob_service_client = get_blob_service_client()
        if not blob_service_client:
            return False

        container_name = get_container_name()
        blob_client = blob_service_client.get_blob_client(container_name, blob_name)
        blob_client.delete_blob()

        logging.info(f"Deleted blob: {blob_name}")
        return True

    except Exception as e:
        logging.error(f"Error deleting blob: {str(e)}")
        return False


def get_blob_storage_status() -> dict:
    """
    Get the current status of blob storage configuration.
    Useful for diagnostics.

    Returns:
        Dictionary with configuration status
    """
    return {
        "sdk_available": BLOB_STORAGE_AVAILABLE,
        "configured": is_blob_storage_configured(),
        "container_name": get_container_name() if is_blob_storage_configured() else None,
        "sas_expiry_hours": get_sas_expiry_hours(),
        "account_name": os.environ.get('AZURE_STORAGE_ACCOUNT_NAME', '(not set)') if not os.environ.get('AZURE_STORAGE_CONNECTION_STRING') else '(from connection string)'
    }

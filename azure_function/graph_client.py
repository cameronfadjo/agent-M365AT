"""
Microsoft Graph API client module for Refresh agent.

This module centralizes all Microsoft Graph API interactions and OAuth 2.0
On-Behalf-Of (OBO) token exchange for the Refresh agent. It provides functions
to authenticate with Microsoft Entra ID, search OneDrive, retrieve file content
and metadata, and upload files to OneDrive.

All HTTP calls to the Graph API are performed with a 30-second timeout and
include proper error handling with descriptive exception messages.
"""

import os
import logging
import json
import requests
from msal import ConfidentialClientApplication

# Module constants
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
HTTP_TIMEOUT = 30

logger = logging.getLogger(__name__)


def exchange_token(user_assertion: str) -> str:
    """
    Exchange a user assertion for an access token using OAuth 2.0 On-Behalf-Of flow.

    This function uses MSAL ConfidentialClientApplication to perform token
    exchange on behalf of the authenticated user.

    Args:
        user_assertion: The user's assertion token (typically from Azure Function auth).

    Returns:
        str: The access token for Microsoft Graph API calls.

    Raises:
        Exception: If token exchange fails, includes error_description from MSAL.
    """
    tenant_id = os.environ.get("ENTRA_TENANT_ID")
    client_id = os.environ.get("ENTRA_CLIENT_ID")
    client_secret = os.environ.get("ENTRA_CLIENT_SECRET")

    if not all([tenant_id, client_id, client_secret]):
        raise Exception(
            "Missing required environment variables: ENTRA_TENANT_ID, "
            "ENTRA_CLIENT_ID, or ENTRA_CLIENT_SECRET"
        )

    authority = f"https://login.microsoftonline.com/{tenant_id}"

    app = ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority,
    )

    logger.info("Attempting OAuth 2.0 On-Behalf-Of token exchange")

    try:
        result = app.acquire_token_on_behalf_of(
            user_assertion=user_assertion,
            scopes=["https://graph.microsoft.com/.default"],
        )

        if "access_token" in result:
            logger.info("Token exchange successful")
            return result["access_token"]
        else:
            error_description = result.get(
                "error_description", "Unknown error during token exchange"
            )
            logger.error(f"Token exchange failed: {error_description}")
            raise Exception(f"Token exchange failed: {error_description}")

    except Exception as e:
        logger.error(f"Exception during token exchange: {str(e)}")
        raise


def extract_token_from_header(req) -> str:
    """
    Extract Bearer token from the Authorization header.

    Args:
        req: An azure.functions.HttpRequest object.

    Returns:
        str: The Bearer token (without "Bearer " prefix).

    Raises:
        ValueError: If the Authorization header is missing or malformed.
    """
    auth_header = req.headers.get("Authorization", "").strip()

    if not auth_header:
        logger.error("Authorization header is missing")
        raise ValueError("Authorization header is missing")

    if not auth_header.startswith("Bearer "):
        logger.error("Authorization header does not start with 'Bearer '")
        raise ValueError("Authorization header must start with 'Bearer '")

    token = auth_header[7:]  # Remove "Bearer " prefix

    if not token:
        logger.error("Authorization header contains empty token")
        raise ValueError("Authorization header contains empty token")

    return token


def search_onedrive(access_token: str, query: str, limit: int = 25) -> list:
    """
    Search for files and folders in OneDrive using the Microsoft Graph search API.

    Args:
        access_token: The access token for Graph API calls.
        query: The search query string.
        limit: Maximum number of results to return (default: 25).

    Returns:
        list: A list of dicts with keys: id, name, path, webUrl, lastModified,
              createdDateTime, size. Returns empty list if no results.

    Raises:
        Exception: On non-200 HTTP response, includes status code and response text.
    """
    url = f"{GRAPH_BASE_URL}/me/drive/root/search(q='{query}')"

    params = {
        "$top": limit,
        "$select": "id,name,parentReference,webUrl,lastModifiedDateTime,createdDateTime,size",
    }

    headers = {"Authorization": f"Bearer {access_token}"}

    logger.info(f"Searching OneDrive with query: {query}")

    try:
        response = requests.get(url, headers=headers, params=params, timeout=HTTP_TIMEOUT)

        if response.status_code != 200:
            error_msg = f"Search failed with status {response.status_code}: {response.text}"
            logger.error(error_msg)
            raise Exception(error_msg)

        data = response.json()
        results = []

        for item in data.get("value", []):
            parent_ref = item.get("parentReference", {})
            result_dict = {
                "id": item.get("id"),
                "name": item.get("name"),
                "path": parent_ref.get("path", ""),
                "webUrl": item.get("webUrl"),
                "lastModified": item.get("lastModifiedDateTime"),
                "createdDateTime": item.get("createdDateTime"),
                "size": item.get("size", 0),
            }
            results.append(result_dict)

        logger.info(f"Search returned {len(results)} results")
        return results

    except requests.RequestException as e:
        error_msg = f"Search request failed: {str(e)}"
        logger.error(error_msg)
        raise Exception(error_msg)


def get_file_content(access_token: str, item_id: str) -> tuple:
    """
    Retrieve the binary content of a file from OneDrive.

    Args:
        access_token: The access token for Graph API calls.
        item_id: The OneDrive item ID of the file.

    Returns:
        tuple: (file_bytes: bytes, content_type: str)

    Raises:
        Exception: On failure, includes status code and error details.
    """
    url = f"{GRAPH_BASE_URL}/me/drive/items/{item_id}/content"

    headers = {"Authorization": f"Bearer {access_token}"}

    logger.info(f"Retrieving file content for item_id: {item_id}")

    try:
        response = requests.get(
            url, headers=headers, timeout=HTTP_TIMEOUT, allow_redirects=True
        )

        if response.status_code != 200:
            error_msg = (
                f"Failed to get file content, status {response.status_code}: "
                f"{response.text}"
            )
            logger.error(error_msg)
            raise Exception(error_msg)

        content_type = response.headers.get("Content-Type", "application/octet-stream")

        logger.info(f"File content retrieved successfully, size: {len(response.content)} bytes")
        return (response.content, content_type)

    except requests.RequestException as e:
        error_msg = f"File content retrieval request failed: {str(e)}"
        logger.error(error_msg)
        raise Exception(error_msg)


def get_file_metadata(access_token: str, item_id: str) -> dict:
    """
    Retrieve metadata for a file in OneDrive.

    Args:
        access_token: The access token for Graph API calls.
        item_id: The OneDrive item ID of the file.

    Returns:
        dict: Contains keys: id, name, createdDateTime, lastModifiedDateTime, size.

    Raises:
        Exception: On failure, includes status code and error details.
    """
    url = f"{GRAPH_BASE_URL}/me/drive/items/{item_id}"

    params = {"$select": "id,name,createdDateTime,lastModifiedDateTime,size,file"}

    headers = {"Authorization": f"Bearer {access_token}"}

    logger.info(f"Retrieving file metadata for item_id: {item_id}")

    try:
        response = requests.get(url, headers=headers, params=params, timeout=HTTP_TIMEOUT)

        if response.status_code != 200:
            error_msg = (
                f"Failed to get file metadata, status {response.status_code}: "
                f"{response.text}"
            )
            logger.error(error_msg)
            raise Exception(error_msg)

        data = response.json()

        metadata = {
            "id": data.get("id"),
            "name": data.get("name"),
            "createdDateTime": data.get("createdDateTime"),
            "lastModifiedDateTime": data.get("lastModifiedDateTime"),
            "size": data.get("size", 0),
        }

        logger.info(f"File metadata retrieved successfully")
        return metadata

    except requests.RequestException as e:
        error_msg = f"File metadata retrieval request failed: {str(e)}"
        logger.error(error_msg)
        raise Exception(error_msg)


def save_file_to_onedrive(
    access_token: str, file_bytes: bytes, filename: str, folder: str = "Refresh"
) -> dict:
    """
    Upload a file to OneDrive, creating the target folder if it doesn't exist.

    This function first checks if the target folder exists. If not, it creates
    the folder with automatic conflict handling. Then it uploads the file to
    the specified folder.

    Args:
        access_token: The access token for Graph API calls.
        file_bytes: The binary content of the file to upload.
        filename: The name to give the uploaded file.
        folder: The folder path where the file should be saved (default: "Refresh").

    Returns:
        dict: Contains keys: success (bool), webUrl (str), itemId (str).

    Raises:
        Exception: On failure, includes status code and error details.
    """
    logger.info(f"Starting file upload to folder '{folder}' with filename '{filename}'")

    # Step 1: Check if folder exists
    folder_check_url = f"{GRAPH_BASE_URL}/me/drive/root:/{folder}"
    headers = {"Authorization": f"Bearer {access_token}"}

    try:
        folder_response = requests.get(
            folder_check_url, headers=headers, timeout=HTTP_TIMEOUT
        )

        if folder_response.status_code == 404:
            # Step 2: Create folder if it doesn't exist
            logger.info(f"Folder '{folder}' does not exist, creating it")

            create_folder_url = f"{GRAPH_BASE_URL}/me/drive/root/children"

            folder_payload = {
                "name": folder,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "rename",
            }

            create_response = requests.post(
                create_folder_url,
                headers=headers,
                json=folder_payload,
                timeout=HTTP_TIMEOUT,
            )

            if create_response.status_code not in (200, 201):
                error_msg = (
                    f"Failed to create folder, status {create_response.status_code}: "
                    f"{create_response.text}"
                )
                logger.error(error_msg)
                raise Exception(error_msg)

            logger.info(f"Folder '{folder}' created successfully")

        elif folder_response.status_code != 200:
            error_msg = (
                f"Failed to check folder, status {folder_response.status_code}: "
                f"{folder_response.text}"
            )
            logger.error(error_msg)
            raise Exception(error_msg)
        else:
            logger.info(f"Folder '{folder}' already exists")

        # Step 3: Upload file to folder
        upload_url = f"{GRAPH_BASE_URL}/me/drive/root:/{folder}/{filename}:/content"

        upload_headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/octet-stream",
        }

        logger.info(f"Uploading file to {upload_url}")

        upload_response = requests.put(
            upload_url, headers=upload_headers, data=file_bytes, timeout=HTTP_TIMEOUT
        )

        if upload_response.status_code not in (200, 201):
            error_msg = (
                f"Failed to upload file, status {upload_response.status_code}: "
                f"{upload_response.text}"
            )
            logger.error(error_msg)
            raise Exception(error_msg)

        upload_data = upload_response.json()

        result = {
            "success": True,
            "webUrl": upload_data.get("webUrl"),
            "itemId": upload_data.get("id"),
        }

        logger.info(f"File uploaded successfully: {result['webUrl']}")
        return result

    except requests.RequestException as e:
        error_msg = f"File upload request failed: {str(e)}"
        logger.error(error_msg)
        raise Exception(error_msg)

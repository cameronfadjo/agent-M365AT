"""
Refresh v6 - Azure Functions (M365 Agents Toolkit Edition)
Agentic document research and generation API.

M365AT endpoints (new — replace Power Automate flows):
- POST /api/search-onedrive        - Search user's OneDrive via Graph API
- POST /api/retrieve-and-analyze   - Fetch docs from OneDrive + comparative analysis
- POST /api/save-to-onedrive       - Save generated doc to user's OneDrive

v6 endpoints:
- POST /api/extract-search-intent  - Simplified search-only intent extraction
- POST /api/analyze-family         - Cross-document comparative analysis
- POST /api/generate-from-synthesis - Generate document from comparative analysis

v5 endpoints (retained for backward compatibility / single-doc fallback):
- POST /api/extract-intent     - Extract intent from natural language (with field extraction)
- POST /api/analyze-document   - Single-document analysis
- POST /api/generate-document  - Generate new document from fields
- POST /api/merge-fields       - Merge user changes into analyzed fields
- POST /api/refresh-document   - Combined analyze + generate workflow

Utility endpoints:
- GET  /api/health             - Health check
- GET  /api/storage-status     - Blob storage status
"""

import azure.functions as func
import json
import logging
import os
import base64
import requests
from datetime import datetime

# Local modules
from document_generator import (
    DocumentGenerator, generate_document, generate_filename,
    generate_from_synthesis
)
from document_analyzer import analyze_document, generate_results_card, generate_input_card
from intent_extractor import (
    extract_intent, extract_search_intent,
    merge_fields, generate_field_input_card
)
from family_analyzer import analyze_document_family
from blob_storage import (
    is_blob_storage_configured,
    upload_document_and_get_sas_url,
    get_blob_storage_status
)
import graph_client

app = func.FunctionApp()


# =============================================================================
# M365 AGENTS TOOLKIT ENDPOINTS (replace Power Automate flows)
# =============================================================================

@app.route(route="search-onedrive", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
def search_onedrive_endpoint(req: func.HttpRequest) -> func.HttpResponse:
    """
    Search user's OneDrive for documents matching search terms.
    Replaces Power Automate Flow 2. Uses OBO token exchange for Graph API.

    Request JSON:
    {
        "search_terms": "back to school letter"
    }
    Authorization header: Bearer <SSO token from M365 Copilot>

    Response JSON:
    {
        "success": true,
        "documents": [
            {
                "id": "...",
                "name": "Back to School Letter 2025.docx",
                "path": "/Documents/",
                "webUrl": "https://...",
                "lastModified": "2025-08-14T10:30:00Z",
                "createdDateTime": "2025-08-10T09:00:00Z",
                "size": 45632
            }
        ],
        "count": 3
    }
    """
    logging.info('OneDrive search request received')

    try:
        # Extract and exchange token
        user_token = graph_client.extract_token_from_header(req)
        access_token = graph_client.exchange_token(user_token)

        # Get search terms
        req_body = req.get_json()
        search_terms = req_body.get('search_terms', '')

        if not search_terms:
            return func.HttpResponse(
                json.dumps({'success': False, 'error': 'Missing search_terms'}),
                status_code=400,
                mimetype='application/json'
            )

        # Search OneDrive
        documents = graph_client.search_onedrive(access_token, search_terms)

        return func.HttpResponse(
            json.dumps({
                'success': True,
                'documents': documents,
                'count': len(documents)
            }),
            status_code=200,
            mimetype='application/json',
            headers={'Access-Control-Allow-Origin': '*'}
        )

    except ValueError as e:
        logging.error(f'Auth error in search-onedrive: {str(e)}')
        return func.HttpResponse(
            json.dumps({'success': False, 'error': f'Authentication error: {str(e)}'}),
            status_code=401,
            mimetype='application/json'
        )
    except Exception as e:
        logging.error(f'Error searching OneDrive: {str(e)}')
        return func.HttpResponse(
            json.dumps({'success': False, 'error': str(e)}),
            status_code=500,
            mimetype='application/json'
        )


@app.route(route="retrieve-and-analyze", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
def retrieve_and_analyze_endpoint(req: func.HttpRequest) -> func.HttpResponse:
    """
    Retrieve documents from OneDrive and run cross-document comparative analysis.
    Replaces Power Automate Flow 3: fetches file contents via Graph API,
    base64-encodes them, and calls analyze_document_family() internally.

    Request JSON:
    {
        "document_ids": ["id1", "id2", "id3"],
        "context_document_ids": ["id4", "id5"],
        "user_context": "I need this year's back-to-school letter"
    }
    Authorization header: Bearer <SSO token from M365 Copilot>

    Response JSON:
    {
        "success": true,
        "family_type": "back_to_school_letter",
        "family_type_display": "Back-to-School Letter",
        "document_count": 3,
        "date_range": "2023–2025",
        "analysis": {...},
        "recommended_base": "Back to School Welcome 2025.docx",
        "base_document_text": "Full text of recommended base...",
        "organizational_context": "New assistant principal Dr. Johnson hired...",
        "confidence": 0.9,
        "summary": "..."
    }
    """
    logging.info('Retrieve-and-analyze request received')

    try:
        # Extract and exchange token
        user_token = graph_client.extract_token_from_header(req)
        access_token = graph_client.exchange_token(user_token)

        # Parse request
        req_body = req.get_json()
        doc_ids = req_body.get('document_ids', [])
        context_ids = req_body.get('context_document_ids', [])
        user_context = req_body.get('user_context', '')

        if not doc_ids:
            return func.HttpResponse(
                json.dumps({'success': False, 'error': 'Missing document_ids'}),
                status_code=400,
                mimetype='application/json'
            )

        # Retrieve document family files from OneDrive
        documents = []
        for doc_id in doc_ids:
            try:
                content_bytes, content_type = graph_client.get_file_content(access_token, doc_id)
                metadata = graph_client.get_file_metadata(access_token, doc_id)
                documents.append({
                    'filename': metadata.get('name', f'document_{doc_id}.docx'),
                    'content': base64.b64encode(content_bytes).decode('utf-8'),
                    'metadata': {
                        'created': metadata.get('createdDateTime', ''),
                        'modified': metadata.get('lastModifiedDateTime', '')
                    }
                })
                logging.info(f'Retrieved document: {metadata.get("name", doc_id)}')
            except Exception as e:
                logging.error(f'Error retrieving document {doc_id}: {str(e)}')

        if not documents:
            return func.HttpResponse(
                json.dumps({'success': False, 'error': 'Could not retrieve any documents'}),
                status_code=400,
                mimetype='application/json'
            )

        # Retrieve context documents from OneDrive (optional)
        context_documents = None
        if context_ids:
            context_documents = []
            for ctx_id in context_ids:
                try:
                    content_bytes, content_type = graph_client.get_file_content(access_token, ctx_id)
                    metadata = graph_client.get_file_metadata(access_token, ctx_id)
                    context_documents.append({
                        'filename': metadata.get('name', f'context_{ctx_id}.docx'),
                        'content': base64.b64encode(content_bytes).decode('utf-8'),
                        'metadata': {
                            'created': metadata.get('createdDateTime', ''),
                            'modified': metadata.get('lastModifiedDateTime', '')
                        }
                    })
                    logging.info(f'Retrieved context document: {metadata.get("name", ctx_id)}')
                except Exception as e:
                    logging.warning(f'Error retrieving context document {ctx_id}: {str(e)}')

        # Run family analysis (reuses existing v6 logic)
        result = analyze_document_family(
            documents, user_context,
            context_documents=context_documents if context_documents else None
        )

        return func.HttpResponse(
            json.dumps({
                'success': True,
                **result
            }),
            status_code=200,
            mimetype='application/json',
            headers={'Access-Control-Allow-Origin': '*'}
        )

    except ValueError as e:
        logging.error(f'Auth error in retrieve-and-analyze: {str(e)}')
        return func.HttpResponse(
            json.dumps({'success': False, 'error': f'Authentication error: {str(e)}'}),
            status_code=401,
            mimetype='application/json'
        )
    except Exception as e:
        logging.error(f'Error in retrieve-and-analyze: {str(e)}')
        return func.HttpResponse(
            json.dumps({'success': False, 'error': str(e)}),
            status_code=500,
            mimetype='application/json'
        )


@app.route(route="save-to-onedrive", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
def save_to_onedrive_endpoint(req: func.HttpRequest) -> func.HttpResponse:
    """
    Download generated document from blob storage and save to user's OneDrive.
    Replaces Power Automate Flow 5.

    Request JSON:
    {
        "download_url": "https://storageaccount.blob.core.windows.net/...?sv=...",
        "filename": "Back to School Letter - 2026-2027.docx",
        "folder_path": "Refresh"
    }
    Authorization header: Bearer <SSO token from M365 Copilot>

    Response JSON:
    {
        "success": true,
        "savedPath": "/Refresh/Back to School Letter - 2026-2027.docx",
        "webUrl": "https://contoso-my.sharepoint.com/...",
        "itemId": "..."
    }
    """
    logging.info('Save-to-OneDrive request received')

    try:
        # Extract and exchange token
        user_token = graph_client.extract_token_from_header(req)
        access_token = graph_client.exchange_token(user_token)

        # Parse request
        req_body = req.get_json()
        download_url = req_body.get('download_url', '')
        filename = req_body.get('filename', '')
        folder_path = req_body.get('folder_path', 'Refresh')

        if not download_url or not filename:
            return func.HttpResponse(
                json.dumps({'success': False, 'error': 'Missing download_url or filename'}),
                status_code=400,
                mimetype='application/json'
            )

        # Download from blob storage SAS URL
        logging.info(f'Downloading generated document from blob storage')
        download_response = requests.get(download_url, timeout=60)
        download_response.raise_for_status()
        file_bytes = download_response.content
        logging.info(f'Downloaded {len(file_bytes)} bytes')

        # Save to user's OneDrive
        save_result = graph_client.save_file_to_onedrive(
            access_token, file_bytes, filename, folder_path
        )

        return func.HttpResponse(
            json.dumps({
                'success': True,
                'savedPath': f'/{folder_path}/{filename}',
                'webUrl': save_result.get('webUrl', ''),
                'itemId': save_result.get('itemId', '')
            }),
            status_code=200,
            mimetype='application/json',
            headers={'Access-Control-Allow-Origin': '*'}
        )

    except ValueError as e:
        logging.error(f'Auth error in save-to-onedrive: {str(e)}')
        return func.HttpResponse(
            json.dumps({'success': False, 'error': f'Authentication error: {str(e)}'}),
            status_code=401,
            mimetype='application/json'
        )
    except requests.RequestException as e:
        logging.error(f'Error downloading from blob storage: {str(e)}')
        return func.HttpResponse(
            json.dumps({'success': False, 'error': f'Download failed: {str(e)}'}),
            status_code=502,
            mimetype='application/json'
        )
    except Exception as e:
        logging.error(f'Error saving to OneDrive: {str(e)}')
        return func.HttpResponse(
            json.dumps({'success': False, 'error': str(e)}),
            status_code=500,
            mimetype='application/json'
        )


# =============================================================================
# v6 ENDPOINTS
# =============================================================================

@app.route(route="extract-search-intent", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
def extract_search_intent_endpoint(req: func.HttpRequest) -> func.HttpResponse:
    """
    Extract search intent only — no field extraction.
    Simplified for the v6 pipeline where searching is separated from editing.

    Request JSON:
    {
        "prompt": "I need this year's back-to-school letter"
    }

    Response JSON:
    {
        "success": true,
        "document_type": "back_to_school_letter",
        "search_terms": ["back to school", "letter", "welcome"],
        "context_search_terms": ["new staff 2026", "budget update", "technology initiative"],
        "summary": "User needs a back-to-school letter for this year",
        "confidence": 0.9
    }
    """
    logging.info('Search intent extraction request received')

    try:
        req_body = req.get_json()
        prompt = req_body.get('prompt', '')

        if not prompt:
            return func.HttpResponse(
                json.dumps({'success': False, 'error': 'Missing prompt field'}),
                status_code=400,
                mimetype='application/json'
            )

        result = extract_search_intent(prompt)

        return func.HttpResponse(
            json.dumps({
                'success': True,
                **result
            }),
            status_code=200,
            mimetype='application/json',
            headers={'Access-Control-Allow-Origin': '*'}
        )

    except Exception as e:
        logging.error(f'Error extracting search intent: {str(e)}')
        return func.HttpResponse(
            json.dumps({'success': False, 'error': str(e)}),
            status_code=500,
            mimetype='application/json'
        )


@app.route(route="analyze-family", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
def analyze_family_endpoint(req: func.HttpRequest) -> func.HttpResponse:
    """
    Cross-document comparative analysis. Core intelligence of v6.

    Accepts multiple documents (base64-encoded) and performs comparative
    analysis to identify stable elements, variable elements, and emerging
    elements across the document family.

    Request JSON:
    {
        "documents": [
            {
                "filename": "Back to School Letter 2023.docx",
                "content": "base64-encoded-content",
                "metadata": {"created": "2023-08-15", "modified": "2023-08-20"}
            },
            ...
        ],
        "context_documents": [
            {
                "filename": "New Hire Announcement.docx",
                "content": "base64-encoded-content",
                "metadata": {"created": "2026-01-15", "modified": "2026-01-15"}
            },
            ...
        ],
        "user_context": "I need this year's back-to-school letter"
    }

    Response JSON:
    {
        "success": true,
        "family_type": "back_to_school_letter",
        "family_type_display": "Back-to-School Letter",
        "document_count": 3,
        "date_range": "2023–2025",
        "analysis": {
            "stable_elements": {...},
            "variable_elements": {...},
            "emerging_elements": {...}
        },
        "recommended_base": "Back to School Welcome 2025.docx",
        "base_document_text": "Full text of recommended base...",
        "organizational_context": "New assistant principal Dr. Johnson hired; 1:1 device program expanding",
        "confidence": 0.9,
        "summary": "..."
    }
    """
    logging.info('Family analysis request received')

    try:
        req_body = req.get_json()
        documents = req_body.get('documents', [])
        context_documents = req_body.get('context_documents', [])
        user_context = req_body.get('user_context', '')

        if not documents:
            return func.HttpResponse(
                json.dumps({'success': False, 'error': 'Missing documents array'}),
                status_code=400,
                mimetype='application/json'
            )

        result = analyze_document_family(
            documents, user_context,
            context_documents=context_documents if context_documents else None
        )

        return func.HttpResponse(
            json.dumps({
                'success': True,
                **result
            }),
            status_code=200,
            mimetype='application/json',
            headers={'Access-Control-Allow-Origin': '*'}
        )

    except ValueError as e:
        return func.HttpResponse(
            json.dumps({'success': False, 'error': str(e)}),
            status_code=400,
            mimetype='application/json'
        )
    except Exception as e:
        logging.error(f'Error analyzing document family: {str(e)}')
        return func.HttpResponse(
            json.dumps({'success': False, 'error': str(e)}),
            status_code=500,
            mimetype='application/json'
        )


@app.route(route="generate-from-synthesis", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
def generate_from_synthesis_endpoint(req: func.HttpRequest) -> func.HttpResponse:
    """
    Generate a new document grounded in comparative analysis.

    Takes the family analysis, the base document text, user changes,
    and target year, then generates a complete new version.

    Request JSON:
    {
        "family_analysis": {...},
        "base_document_text": "The full text of the most recent version",
        "organizational_context": "New assistant principal Dr. Johnson; 1:1 device program",
        "user_changes": "Change the principal to Dr. Johnson",
        "target_year": "2026-2027"
    }

    Response JSON:
    {
        "success": true,
        "generated_text": "The complete document text...",
        "changes_applied": ["Updated school year to 2026-2027", ...],
        "flags": [{"field": "first_day_of_school", "reason": "...", "placeholder": "[...]"}],
        "filename": "Back to School Letter - 2026-2027.docx",
        "download_url": "https://...",
        "expires_in_hours": 24
    }
    """
    logging.info('Synthesis generation request received')

    try:
        req_body = req.get_json()
        family_analysis = req_body.get('family_analysis', {})
        base_text = req_body.get('base_document_text', '')
        organizational_context = req_body.get('organizational_context', '')
        user_changes = req_body.get('user_changes', '')
        target_year = req_body.get('target_year', '')

        if not family_analysis or not base_text:
            return func.HttpResponse(
                json.dumps({'success': False, 'error': 'Missing family_analysis or base_document_text'}),
                status_code=400,
                mimetype='application/json'
            )

        # Generate text from synthesis
        synthesis_result = generate_from_synthesis(
            family_analysis, base_text, user_changes, target_year,
            organizational_context=organizational_context
        )

        # Create Word document from generated text
        doc_type = family_analysis.get('family_type', 'letter')
        generated_text = synthesis_result.get('generated_text', '')
        filename = synthesis_result.get(
            'suggested_filename',
            f'Generated Document - {target_year or "New"}.docx'
        )

        # Build fields for document generator
        fields = {'body': generated_text}
        doc_bytes = generate_document(doc_type, fields)

        # Upload to blob storage
        if is_blob_storage_configured():
            success, sas_url, _ = upload_document_and_get_sas_url(doc_bytes, filename)

            if success:
                return func.HttpResponse(
                    json.dumps({
                        'success': True,
                        'generated_text': generated_text,
                        'changes_applied': synthesis_result.get('changes_applied', []),
                        'flags': synthesis_result.get('flags', []),
                        'filename': filename,
                        'download_url': sas_url,
                        'expires_in_hours': 24,
                        'storage_type': 'blob_sas_url'
                    }),
                    status_code=200,
                    mimetype='application/json',
                    headers={'Access-Control-Allow-Origin': '*'}
                )
            else:
                logging.error(f'Blob upload failed for synthesis: {sas_url}')

        # Fallback to base64
        return func.HttpResponse(
            json.dumps({
                'success': True,
                'generated_text': generated_text,
                'changes_applied': synthesis_result.get('changes_applied', []),
                'flags': synthesis_result.get('flags', []),
                'filename': filename,
                'content': base64.b64encode(doc_bytes).decode('utf-8'),
                'storage_type': 'base64'
            }),
            status_code=200,
            mimetype='application/json',
            headers={'Access-Control-Allow-Origin': '*'}
        )

    except ValueError as e:
        return func.HttpResponse(
            json.dumps({'success': False, 'error': str(e)}),
            status_code=400,
            mimetype='application/json'
        )
    except Exception as e:
        logging.error(f'Error generating from synthesis: {str(e)}')
        return func.HttpResponse(
            json.dumps({'success': False, 'error': str(e)}),
            status_code=500,
            mimetype='application/json'
        )


# =============================================================================
# v5 ENDPOINTS (retained for backward compatibility / single-doc fallback)
# =============================================================================

@app.route(route="extract-intent", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
def extract_intent_endpoint(req: func.HttpRequest) -> func.HttpResponse:
    """
    Extract user intent and field values from natural language.
    v5 endpoint — kept for backward compatibility and single-doc fallback.
    """
    logging.info('Intent extraction request received')

    try:
        req_body = req.get_json()
        prompt = req_body.get('prompt', '')

        if not prompt:
            return func.HttpResponse(
                json.dumps({'success': False, 'error': 'Missing prompt field'}),
                status_code=400,
                mimetype='application/json'
            )

        result = extract_intent(prompt)

        return func.HttpResponse(
            json.dumps({
                'success': True,
                **result
            }),
            status_code=200,
            mimetype='application/json',
            headers={'Access-Control-Allow-Origin': '*'}
        )

    except Exception as e:
        logging.error(f'Error extracting intent: {str(e)}')
        return func.HttpResponse(
            json.dumps({'success': False, 'error': str(e)}),
            status_code=500,
            mimetype='application/json'
        )


@app.route(route="analyze-document", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
def analyze_document_endpoint(req: func.HttpRequest) -> func.HttpResponse:
    """
    Analyze an uploaded document and extract its variable fields.
    v5 endpoint — kept for single-document fallback path.
    """
    logging.info('Document analysis request received')

    try:
        req_body = req.get_json()

        filename = req_body.get('filename', 'document.docx')
        content_b64 = req_body.get('content', '')
        extracted_fields = req_body.get('extracted_fields', {})

        if not content_b64:
            return func.HttpResponse(
                json.dumps({'success': False, 'error': 'Missing content field'}),
                status_code=400,
                mimetype='application/json'
            )

        try:
            file_bytes = base64.b64decode(content_b64)
        except Exception as e:
            return func.HttpResponse(
                json.dumps({'success': False, 'error': f'Invalid base64 content: {str(e)}'}),
                status_code=400,
                mimetype='application/json'
            )

        analysis = analyze_document(file_bytes, filename)

        if extracted_fields:
            analysis['fields'] = merge_fields(analysis.get('fields', []), extracted_fields)

        analysis['results_card'] = generate_results_card(analysis)
        analysis['input_card'] = generate_input_card(analysis)

        return func.HttpResponse(
            json.dumps({
                'success': True,
                **analysis
            }),
            status_code=200,
            mimetype='application/json',
            headers={'Access-Control-Allow-Origin': '*'}
        )

    except ValueError as e:
        return func.HttpResponse(
            json.dumps({'success': False, 'error': str(e)}),
            status_code=400,
            mimetype='application/json'
        )
    except Exception as e:
        logging.error(f'Error analyzing document: {str(e)}')
        return func.HttpResponse(
            json.dumps({'success': False, 'error': str(e)}),
            status_code=500,
            mimetype='application/json'
        )


@app.route(route="generate-document", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
def generate_document_endpoint(req: func.HttpRequest) -> func.HttpResponse:
    """
    Generate a new document with updated fields.
    v5 endpoint — kept for single-document generation.
    """
    logging.info('Document generation request received')

    try:
        req_body = req.get_json()

        document_type = req_body.get('document_type', 'memo')
        fields = req_body.get('fields', {})
        template_b64 = req_body.get('template_content')
        custom_filename = req_body.get('filename')

        if not fields:
            return func.HttpResponse(
                json.dumps({'success': False, 'error': 'Missing fields'}),
                status_code=400,
                mimetype='application/json'
            )

        template_bytes = None
        if template_b64:
            try:
                template_bytes = base64.b64decode(template_b64)
            except Exception as e:
                logging.warning(f'Invalid template base64, ignoring: {e}')

        doc_bytes = generate_document(document_type, fields, template_bytes)
        filename = custom_filename or generate_filename(document_type, fields)

        if is_blob_storage_configured():
            success, message, sas_url = upload_document_and_get_sas_url(doc_bytes, filename)

            if success:
                return func.HttpResponse(
                    json.dumps({
                        'success': True,
                        'filename': filename,
                        'download_url': sas_url,
                        'expires_in_hours': 24,
                        'storage_type': 'blob_sas_url'
                    }),
                    status_code=200,
                    mimetype='application/json',
                    headers={'Access-Control-Allow-Origin': '*'}
                )
            else:
                logging.error(f'Blob upload failed: {message}')
                return func.HttpResponse(
                    json.dumps({
                        'success': True,
                        'filename': filename,
                        'content': base64.b64encode(doc_bytes).decode('utf-8'),
                        'content_type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                        'storage_type': 'base64',
                        'warning': f'Blob storage failed: {message}'
                    }),
                    status_code=200,
                    mimetype='application/json',
                    headers={'Access-Control-Allow-Origin': '*'}
                )
        else:
            return func.HttpResponse(
                json.dumps({
                    'success': True,
                    'filename': filename,
                    'content': base64.b64encode(doc_bytes).decode('utf-8'),
                    'content_type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    'storage_type': 'base64'
                }),
                status_code=200,
                mimetype='application/json',
                headers={'Access-Control-Allow-Origin': '*'}
            )

    except Exception as e:
        logging.error(f'Error generating document: {str(e)}')
        return func.HttpResponse(
            json.dumps({'success': False, 'error': str(e)}),
            status_code=500,
            mimetype='application/json'
        )


@app.route(route="refresh-document", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
def refresh_document_endpoint(req: func.HttpRequest) -> func.HttpResponse:
    """
    Combined endpoint: Analyze document and generate updated version in one call.
    v5 endpoint — kept for simpler integrations.
    """
    logging.info('Refresh document request received')

    try:
        req_body = req.get_json()

        filename = req_body.get('filename', 'document.docx')
        content_b64 = req_body.get('content', '')
        updated_fields = req_body.get('updated_fields', {})
        output_filename = req_body.get('output_filename')

        if not content_b64:
            return func.HttpResponse(
                json.dumps({'success': False, 'error': 'Missing content field'}),
                status_code=400,
                mimetype='application/json'
            )

        file_bytes = base64.b64decode(content_b64)

        analysis = analyze_document(file_bytes, filename)

        merged_fields = {}
        for field in analysis.get('fields', []):
            field_name = field.get('field_name')
            if field_name in updated_fields:
                merged_fields[field_name] = updated_fields[field_name]
            else:
                merged_fields[field_name] = field.get('current_value', '')

        document_type = analysis.get('document_type', 'memo')
        doc_bytes = generate_document(document_type, merged_fields)

        final_filename = output_filename or generate_filename(document_type, merged_fields)

        if is_blob_storage_configured():
            success, message, sas_url = upload_document_and_get_sas_url(doc_bytes, final_filename)

            if success:
                return func.HttpResponse(
                    json.dumps({
                        'success': True,
                        'analysis': {
                            'document_type': analysis.get('document_type'),
                            'document_type_display': analysis.get('document_type_display'),
                            'confidence': analysis.get('confidence'),
                            'summary': analysis.get('summary'),
                            'field_count': len(analysis.get('fields', []))
                        },
                        'generated': {
                            'filename': final_filename,
                            'download_url': sas_url,
                            'expires_in_hours': 24
                        }
                    }),
                    status_code=200,
                    mimetype='application/json',
                    headers={'Access-Control-Allow-Origin': '*'}
                )

        return func.HttpResponse(
            json.dumps({
                'success': True,
                'analysis': {
                    'document_type': analysis.get('document_type'),
                    'summary': analysis.get('summary')
                },
                'generated': {
                    'filename': final_filename,
                    'content': base64.b64encode(doc_bytes).decode('utf-8')
                }
            }),
            status_code=200,
            mimetype='application/json',
            headers={'Access-Control-Allow-Origin': '*'}
        )

    except Exception as e:
        logging.error(f'Error refreshing document: {str(e)}')
        return func.HttpResponse(
            json.dumps({'success': False, 'error': str(e)}),
            status_code=500,
            mimetype='application/json'
        )


# =============================================================================
# MERGE FIELDS HELPERS
# =============================================================================

def _parse_natural_language_changes(
    user_text: str,
    original_fields: list
) -> dict:
    """
    Use Azure OpenAI to parse natural language change requests into
    a structured dict of field_name -> new_value.
    """
    endpoint = os.environ.get('AZURE_OPENAI_ENDPOINT')
    api_key = os.environ.get('AZURE_OPENAI_KEY')
    deployment = os.environ.get('AZURE_OPENAI_DEPLOYMENT', 'gpt-4o-mini')
    api_version = os.environ.get('AZURE_OPENAI_API_VERSION', '2025-01-01-preview')

    if not endpoint or not api_key:
        logging.warning('Azure OpenAI not configured — cannot parse natural language changes')
        return {}

    field_list = []
    for f in original_fields:
        field_list.append(
            f"- {f.get('field_name', '?')} (label: {f.get('field_label', '?')}, "
            f"current: {f.get('current_value', '?')})"
        )
    field_reference = "\n".join(field_list)

    system_prompt = f"""You are a field-change parser. Given a user's natural language request and a list of available document fields, extract ONLY the specific field changes the user wants to make.

Available fields in this document:
{field_reference}

Rules:
1. Match the user's request to the closest field_name from the list above.
2. Return ONLY fields the user explicitly wants to change.
3. Use the exact field_name from the list (not the label).
4. Extract the new value the user wants — use their exact wording for the value.
5. If the user's request doesn't clearly map to any field, return an empty object.

Return ONLY valid JSON — a flat object mapping field_name to new_value:
{{"field_name": "new value", "another_field": "another value"}}

Examples:
User says: "Change the date to March 1, 2026"
If "date" is in the field list → {{"date": "March 1, 2026"}}

User says: "Everything looks good, generate it"
→ {{}}"""

    try:
        url = f"{endpoint}/openai/deployments/{deployment}/chat/completions?api-version={api_version}"
        headers = {
            'Content-Type': 'application/json',
            'api-key': api_key
        }
        payload = {
            'messages': [
                {'role': 'system', 'content': system_prompt},
                {'role': 'user', 'content': user_text}
            ],
            'temperature': 0.1,
            'max_tokens': 500
        }

        import urllib.request
        req_obj = urllib.request.Request(
            url,
            data=json.dumps(payload).encode('utf-8'),
            headers=headers,
            method='POST'
        )

        with urllib.request.urlopen(req_obj, timeout=30) as response:
            result = json.loads(response.read().decode('utf-8'))

        content = result['choices'][0]['message']['content'].strip()

        if content.startswith('```'):
            content = content.split('\n', 1)[-1]
            if content.endswith('```'):
                content = content[:-3].strip()

        parsed = json.loads(content)
        if isinstance(parsed, dict):
            logging.info(f'Parsed natural language changes: {list(parsed.keys())}')
            return parsed
        else:
            logging.warning(f'LLM returned non-dict: {type(parsed)}')
            return {}

    except Exception as e:
        logging.error(f'Error parsing natural language changes: {str(e)}')
        return {}


# =============================================================================
# MERGE FIELDS ENDPOINT
# =============================================================================

@app.route(route="merge-fields", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
def merge_fields_endpoint(req: func.HttpRequest) -> func.HttpResponse:
    """
    Merge user-provided field updates into the original analyzed fields.
    v5 endpoint — kept for single-document editing workflow.
    """
    logging.info('Merge fields request received')

    try:
        req_body = req.get_json()

        original_fields = req_body.get('original_fields', [])
        user_changes = req_body.get('user_changes', {})
        pre_extracted_fields = req_body.get('pre_extracted_fields', {})

        if not original_fields:
            return func.HttpResponse(
                json.dumps({'success': False, 'error': 'Missing original_fields'}),
                status_code=400,
                mimetype='application/json'
            )

        if isinstance(user_changes, str) and user_changes.strip():
            user_changes = _parse_natural_language_changes(
                user_changes, original_fields
            )
        elif not isinstance(user_changes, dict):
            user_changes = {}

        if isinstance(pre_extracted_fields, str):
            try:
                pre_extracted_fields = json.loads(pre_extracted_fields) if pre_extracted_fields.strip() else {}
            except (json.JSONDecodeError, ValueError):
                pre_extracted_fields = {}

        combined_changes = {**pre_extracted_fields, **user_changes}

        merged_detail = merge_fields(original_fields, combined_changes)

        merged_fields_flat = {}
        changed_fields = []
        for field in merged_detail:
            field_name = field.get('field_name', '')
            new_value = field.get('new_value', field.get('current_value', ''))
            merged_fields_flat[field_name] = new_value

            current = field.get('current_value', '')
            if new_value != current:
                changed_fields.append(field_name)
                field['changed'] = True
            else:
                field['changed'] = False

        total = len(merged_detail)
        changed_count = len(changed_fields)
        if changed_fields:
            summary = f"Updated {changed_count} of {total} fields: {', '.join(changed_fields)}"
        else:
            summary = f"No changes detected across {total} fields"

        return func.HttpResponse(
            json.dumps({
                'success': True,
                'merged_fields': merged_fields_flat,
                'fields_detail': merged_detail,
                'changes_summary': summary
            }),
            status_code=200,
            mimetype='application/json',
            headers={'Access-Control-Allow-Origin': '*'}
        )

    except Exception as e:
        logging.error(f'Error merging fields: {str(e)}')
        return func.HttpResponse(
            json.dumps({'success': False, 'error': str(e)}),
            status_code=500,
            mimetype='application/json'
        )


# =============================================================================
# HEALTH CHECK
# =============================================================================

@app.route(route="health", methods=["GET"], auth_level=func.AuthLevel.ANONYMOUS)
def health_endpoint(req: func.HttpRequest) -> func.HttpResponse:
    """Health check endpoint."""
    openai_configured = bool(
        os.environ.get('AZURE_OPENAI_ENDPOINT') and
        os.environ.get('AZURE_OPENAI_KEY')
    )

    blob_configured = is_blob_storage_configured()

    # Check if large model is configured
    large_model = os.environ.get('AZURE_OPENAI_DEPLOYMENT_LARGE', '')

    return func.HttpResponse(
        json.dumps({
            'status': 'healthy',
            'version': '6.0',
            'timestamp': datetime.utcnow().isoformat(),
            'services': {
                'azure_openai': openai_configured,
                'azure_openai_large_model': bool(large_model),
                'blob_storage': blob_configured
            },
            'endpoints': [
                'POST /api/search-onedrive',
                'POST /api/retrieve-and-analyze',
                'POST /api/save-to-onedrive',
                'POST /api/extract-search-intent',
                'POST /api/analyze-family',
                'POST /api/generate-from-synthesis',
                'POST /api/extract-intent',
                'POST /api/analyze-document',
                'POST /api/generate-document',
                'POST /api/merge-fields',
                'POST /api/refresh-document',
                'GET /api/health',
                'GET /api/storage-status'
            ]
        }),
        status_code=200,
        mimetype='application/json',
        headers={'Access-Control-Allow-Origin': '*'}
    )


# =============================================================================
# STORAGE STATUS
# =============================================================================

@app.route(route="storage-status", methods=["GET"], auth_level=func.AuthLevel.ANONYMOUS)
def storage_status_endpoint(req: func.HttpRequest) -> func.HttpResponse:
    """Check blob storage configuration status."""
    return func.HttpResponse(
        json.dumps(get_blob_storage_status()),
        status_code=200,
        mimetype='application/json',
        headers={'Access-Control-Allow-Origin': '*'}
    )

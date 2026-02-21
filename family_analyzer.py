"""
Family Analyzer Module - Refresh v6
Cross-document comparative analysis engine.

Analyzes multiple versions of the same document type to identify:
- Stable elements (consistent across all versions)
- Variable elements (change predictably, e.g., annually)
- Emerging elements (new additions in recent versions)

This is the core intelligence of v6 — the piece that makes the agent
"agentic" by enabling it to research and reason across documents.
"""

import io
import json
import os
import logging
import base64
from typing import Dict, List, Optional, Any

from openai import AzureOpenAI
from document_analyzer import extract_text


# =============================================================================
# COMPARATIVE ANALYSIS PROMPT
# =============================================================================

FAMILY_ANALYSIS_PROMPT = """You are a document family analyst for a school district.
You will receive the full text of multiple versions of the same type of document
(e.g., back-to-school letters from 2023, 2024, 2025).

Your task is to perform a COMPARATIVE ANALYSIS across all versions.

Analyze and categorize every element into:

1. STABLE ELEMENTS: Things that stay the same across ALL versions
   - Document structure (section order, overall format)
   - Tone and style
   - Boilerplate language, standard closings
   - Recurring phrases that never change

2. VARIABLE ELEMENTS: Things that change between versions
   For each, provide:
   - field_name: standardized snake_case name
   - pattern: how it changes (annually, ad hoc, etc.)
   - values_seen: array of values from each version (oldest to newest)
   - predicted_next: your best prediction for the next version, or null if unpredictable
   Common variables: school_year, dates, principal_name, sender_name, event_dates

3. EMERGING ELEMENTS: Things that appeared in recent versions but not earlier ones
   - New sections, topics, or requirements added over time
   - Note which version introduced them

Return ONLY valid JSON:
{
  "family_type": "back_to_school_letter",
  "family_type_display": "Back-to-School Letter",
  "document_count": 3,
  "date_range": "2023–2025",
  "analysis": {
    "stable_elements": {
      "description": "Elements consistent across all versions",
      "items": [
        {"element": "string", "detail": "string"}
      ]
    },
    "variable_elements": {
      "description": "Elements that change with each version",
      "items": [
        {
          "field_name": "string",
          "pattern": "string",
          "values_seen": ["string"],
          "predicted_next": "string or null"
        }
      ]
    },
    "emerging_elements": {
      "description": "Elements added in recent versions",
      "items": [
        {"element": "string", "first_appeared": "string", "detail": "string"}
      ]
    }
  },
  "recommended_base": "filename of most recent version",
  "confidence": 0.0-1.0,
  "summary": "Brief summary of findings"
}

IMPORTANT:
- Cite specific evidence: when you say a field changes annually, show the values
- If you can predict the next value (e.g., school year increments), do so
- If you cannot predict (e.g., exact date), set predicted_next to null
- Be thorough — extract EVERY variable element, even small ones
- The recommended_base should be the most recent version's filename

If organizational context documents are provided (labeled "ORGANIZATIONAL CONTEXT DOCUMENTS"),
extract a brief summary of changes or updates that are relevant to the target document.
Return this as an "organizational_context" field — a plain-text summary of relevant
organizational changes discovered. For example:
- "New assistant principal Dr. Johnson hired (from hire announcement)"
- "1:1 device program expanding to all grades (from budget memo)"
- "Canvas LMS adoption planned for fall 2026 (from technology plan)"

Separate each item with a semicolon. If no context documents are provided or none are
relevant, return "organizational_context": ""."""


# =============================================================================
# MAIN ANALYSIS FUNCTION
# =============================================================================

def analyze_document_family(
    documents: List[Dict[str, Any]],
    user_context: str = '',
    context_documents: Optional[List[Dict[str, Any]]] = None,
    azure_endpoint: Optional[str] = None,
    azure_api_key: Optional[str] = None,
    azure_deployment: Optional[str] = None,
    api_version: Optional[str] = None
) -> Dict[str, Any]:
    """
    Perform comparative analysis across a family of documents,
    optionally incorporating organizational context documents.

    Args:
        documents: List of dicts, each with:
            - 'filename': str
            - 'content': str (base64-encoded file content)
            - 'metadata': dict with 'created' and 'modified' datetime strings
        user_context: The user's original request for context
        context_documents: Optional list of organizational context documents
            (same format as documents). These are recent memos, announcements,
            plans, etc. that might contain relevant organizational changes.
        azure_endpoint: Azure OpenAI endpoint (or env AZURE_OPENAI_ENDPOINT)
        azure_api_key: Azure OpenAI API key (or env AZURE_OPENAI_KEY)
        azure_deployment: Deployment name (or env var — prefers LARGE model)
        api_version: API version (or env AZURE_OPENAI_API_VERSION)

    Returns:
        Dictionary with comparative analysis results:
        - family_type, family_type_display
        - document_count, date_range
        - analysis (stable_elements, variable_elements, emerging_elements)
        - recommended_base, confidence, summary
        - base_document_text (full text of the recommended base document)
        - organizational_context (summary of relevant org changes, or "")
    """
    # Get Azure OpenAI config
    endpoint = azure_endpoint or os.environ.get('AZURE_OPENAI_ENDPOINT')
    api_key = azure_api_key or os.environ.get('AZURE_OPENAI_KEY')
    # Prefer larger model for comparative analysis (more complex reasoning)
    deployment = azure_deployment or os.environ.get(
        'AZURE_OPENAI_DEPLOYMENT_LARGE',
        os.environ.get('AZURE_OPENAI_DEPLOYMENT', 'gpt-4o-mini')
    )
    version = api_version or os.environ.get('AZURE_OPENAI_API_VERSION', '2025-01-01-preview')

    if not endpoint or not api_key:
        raise ValueError(
            "Azure OpenAI credentials not configured. "
            "Set AZURE_OPENAI_ENDPOINT and AZURE_OPENAI_KEY environment variables."
        )

    # Extract text from each document
    doc_texts = []
    for doc in documents:
        try:
            content_bytes = base64.b64decode(doc['content'])
            text = extract_text(content_bytes, doc['filename'])
            if text.strip():
                # Limit per-document text to manage total token budget
                # 3 docs × 6000 chars ≈ 18,000 chars ≈ 4,500 tokens of input
                doc_texts.append({
                    'filename': doc['filename'],
                    'text': text[:6000],
                    'metadata': doc.get('metadata', {})
                })
            else:
                logging.warning(f"Empty text extracted from {doc.get('filename')}")
        except Exception as e:
            logging.warning(f"Could not extract text from {doc.get('filename')}: {e}")

    if not doc_texts:
        raise ValueError("Could not extract text from any of the provided documents.")

    if len(doc_texts) == 1:
        logging.info("Only one document provided — returning single-document stub")
        return {
            'family_type': 'unknown',
            'family_type_display': 'Single Document',
            'document_count': 1,
            'date_range': doc_texts[0]['metadata'].get('created', 'unknown'),
            'analysis': {
                'stable_elements': {'description': 'N/A for single document', 'items': []},
                'variable_elements': {'description': 'N/A for single document', 'items': []},
                'emerging_elements': {'description': 'N/A for single document', 'items': []},
            },
            'recommended_base': doc_texts[0]['filename'],
            'base_document_text': doc_texts[0]['text'],
            'organizational_context': '',
            'confidence': 0.5,
            'summary': f"Only one document found: {doc_texts[0]['filename']}. Use single-document analysis instead.",
            'single_document_fallback': True
        }

    # Sort by created date (oldest first) for chronological analysis
    doc_texts.sort(key=lambda d: d['metadata'].get('created', ''))

    # Build the comparison prompt
    comparison_text = ""
    for i, dt in enumerate(doc_texts, 1):
        created = dt['metadata'].get('created', 'unknown date')
        comparison_text += f"\n\n=== DOCUMENT {i}: {dt['filename']} (created: {created}) ===\n"
        comparison_text += dt['text']

    if user_context:
        comparison_text += f"\n\nUser context: {user_context}"

    # Extract text from context documents (organizational context)
    context_texts = []
    if context_documents:
        for doc in context_documents:
            try:
                content_bytes = base64.b64decode(doc['content'])
                text = extract_text(content_bytes, doc['filename'])
                if text.strip():
                    context_texts.append({
                        'filename': doc['filename'],
                        'text': text[:4000],  # Shorter limit for context docs
                        'metadata': doc.get('metadata', {})
                    })
            except Exception as e:
                logging.warning(f"Could not extract context doc {doc.get('filename')}: {e}")

    if context_texts:
        comparison_text += "\n\n=== ORGANIZATIONAL CONTEXT DOCUMENTS ==="
        for i, ct in enumerate(context_texts, 1):
            comparison_text += f"\n\n--- Context Doc {i}: {ct['filename']} ---\n"
            comparison_text += ct['text']
        logging.info(f"Including {len(context_texts)} organizational context documents")

    # Call Azure OpenAI
    client = AzureOpenAI(
        azure_endpoint=endpoint,
        api_key=api_key,
        api_version=version
    )

    try:
        response = client.chat.completions.create(
            model=deployment,
            messages=[
                {"role": "system", "content": FAMILY_ANALYSIS_PROMPT},
                {"role": "user", "content": f"Analyze these {len(doc_texts)} related documents:\n{comparison_text}"}
            ],
            temperature=0.2,  # Low but not zero — allow some reasoning flexibility
            max_tokens=3000,
            response_format={"type": "json_object"}
        )

        result_text = response.choices[0].message.content
        result = json.loads(result_text)

        # Ensure document_count reflects actual count
        result['document_count'] = len(doc_texts)

        # Attach the base document's full text so Flow 4 can use it for generation
        recommended = result.get('recommended_base', '')
        base_text = ''
        for dt in doc_texts:
            if dt['filename'] == recommended:
                base_text = dt['text']
                break
        # If no match (e.g., LLM returned a slightly different filename), use the last doc (most recent)
        if not base_text and doc_texts:
            base_text = doc_texts[-1]['text']
        result['base_document_text'] = base_text

        # Ensure organizational_context field exists (LLM should return it, but be safe)
        if 'organizational_context' not in result:
            result['organizational_context'] = ''

        logging.info(
            f"Family analysis complete: {result.get('family_type_display', '?')} "
            f"({len(doc_texts)} docs, {len(context_texts)} context docs, "
            f"confidence: {result.get('confidence', '?')})"
        )

        return result

    except Exception as e:
        logging.error(f"Error in family analysis: {e}")
        raise


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def extract_base_document_text(
    content_b64: str,
    filename: str
) -> str:
    """
    Extract plain text from a base64-encoded document.
    Used to get the recommended base document's text for generation.

    Args:
        content_b64: Base64-encoded file content
        filename: Original filename

    Returns:
        Extracted plain text
    """
    file_bytes = base64.b64decode(content_b64)
    return extract_text(file_bytes, filename)

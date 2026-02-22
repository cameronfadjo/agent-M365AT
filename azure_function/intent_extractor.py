"""
Intent Extractor Module - Refresh v6
Uses Azure OpenAI to extract user intent from natural language.

v6 changes:
- New function: extract_search_intent() — focused on search only (no field extraction)
- Original extract_intent() retained for backward compatibility / single-doc fallback
- Improved search term generation with v5 lessons learned

Example:
    User: "I need this year's back-to-school letter"

    extract_search_intent() returns:
    - document_type: back_to_school_letter
    - search_terms: ["back to school", "letter", "welcome"]
    - summary: "User needs a back-to-school letter for this year"
"""

import json
import os
import logging
from typing import Dict, List, Optional, Any

from openai import AzureOpenAI


# =============================================================================
# SEARCH INTENT PROMPT (NEW for v6 — simplified, search-only)
# =============================================================================

SEARCH_INTENT_PROMPT = """You are a search intent extractor for a school district document agent.
Analyze the user's request and extract what is needed to:
1. Search their OneDrive for versions of the target document (search_terms)
2. Search for recent organizational documents that might contain relevant updates (context_search_terms)

Extract:

1. **document_type**: Type of document mentioned
   - "memo" or "memorandum"
   - "letter"
   - "back_to_school_letter"
   - "survey_request"
   - "permission_slip"
   - "announcement"
   - "policy_update"
   - "report"
   - "unknown"

2. **search_terms**: Keywords to find the TARGET DOCUMENT FAMILY (array of strings)
   - These are KEYWORD SEARCH terms, not natural language
   - Graph API matches against file NAMES and file CONTENT
   - Use words likely to appear in file names or inside the document text
   - DO NOT include filler words: "recent", "my", "some", "the", "latest", "new", "old"
   - DO NOT repeat the same word multiple times
   - Keep each term to 1-3 words maximum
   - Aim for 2-4 search terms total

3. **context_search_terms**: Keywords to find RECENT ORGANIZATIONAL DOCUMENTS that might
   contain updates relevant to this document type (array of strings)
   - Think about what kinds of organizational changes would affect this document
   - For a back-to-school letter: staffing changes, new programs, budget updates, technology initiatives
   - For a budget memo: spending reports, purchase orders, project status, device inventory
   - For an AUP/policy: policy updates, board minutes, compliance changes, AI guidelines
   - Include the current year to find recent documents
   - Aim for 3-6 context search terms
   - These should be DIFFERENT from search_terms — they find different documents

4. **summary**: Brief one-sentence description of what the user wants

5. **confidence**: 0.0 to 1.0 indicating extraction confidence

Return ONLY valid JSON:
{
  "document_type": "string",
  "search_terms": ["string"],
  "context_search_terms": ["string"],
  "summary": "string",
  "confidence": 0.0-1.0
}

Examples:

User: "I need this year's back-to-school letter"
{
  "document_type": "back_to_school_letter",
  "search_terms": ["back to school", "back-to-school", "welcome letter"],
  "context_search_terms": ["new staff 2026", "budget update", "technology initiative", "new programs", "calendar 2026-2027"],
  "summary": "User needs a back-to-school letter for this year",
  "confidence": 0.9
}

User: "Update the budget memo"
{
  "document_type": "memo",
  "search_terms": ["budget", "memo", "budget memo"],
  "context_search_terms": ["spending report", "purchase order", "device inventory", "project status"],
  "summary": "User wants to update a budget memo",
  "confidence": 0.85
}

User: "Generate this year's student AUP acknowledgment form"
{
  "document_type": "policy_update",
  "search_terms": ["AUP", "acceptable use", "acknowledgment"],
  "context_search_terms": ["AI policy", "board minutes", "digital citizenship", "BYOD policy", "student technology"],
  "summary": "User needs an updated student AUP acknowledgment form",
  "confidence": 0.9
}

User: "Show me what documents I have"
{
  "document_type": "unknown",
  "search_terms": ["memo", "letter", "notice", "report"],
  "context_search_terms": [],
  "summary": "User wants to browse their documents broadly",
  "confidence": 0.4
}"""


# =============================================================================
# v5 INTENT EXTRACTION PROMPT (kept for backward compatibility / fallback)
# =============================================================================

INTENT_EXTRACTION_PROMPT = """You are an intent extraction assistant. Your job is to analyze user requests about document updates and extract structured information.

When given a user request, extract:

1. **intent**: What the user wants to do
   - "update_document" - User wants to modify an existing document
   - "find_document" - User wants to search for a document
   - "create_document" - User wants to create a new document
   - "unknown" - Cannot determine intent

2. **document_type**: Type of document mentioned
   - "memo" or "memorandum"
   - "letter"
   - "back_to_school_letter"
   - "survey_request"
   - "permission_slip"
   - "announcement"
   - "policy_update"
   - "unknown"

3. **search_terms**: Keywords to search the user's OneDrive via Microsoft Graph API (array of strings)
   - These are KEYWORD SEARCH terms, not natural language
   - Use words likely to appear in file names or inside the document text
   - DO NOT include filler words like "recent", "my", "some", "the", "latest", "new", "old"
   - DO NOT repeat the same word multiple times
   - DO include: document type words ("memo", "letter", "notice"), topic keywords ("budget", "training", "back to school"), and names/dates if mentioned
   - For vague/browsing requests, use BROAD document type terms
   - For specific requests, use SPECIFIC terms
   - Keep each search term to 1-3 words maximum
   - Aim for 2-4 search terms total

4. **extracted_fields**: Any field values the user explicitly mentioned they want to change
   - Map field names to their new values
   - Common fields: date, recipient, sender, principal_name, subject, etc.

5. **confidence**: 0.0 to 1.0 indicating confidence in extraction

Return ONLY valid JSON in this exact format:
{
  "intent": "update_document|find_document|create_document|unknown",
  "document_type": "string",
  "search_terms": ["string", "string"],
  "extracted_fields": {
    "field_name": "value"
  },
  "confidence": 0.0-1.0,
  "summary": "Brief one-sentence summary of what the user wants"
}

Examples:

User: "I need to update the back-to-school letter with the new principal Dr. Johnson"
{
  "intent": "update_document",
  "document_type": "back_to_school_letter",
  "search_terms": ["back-to-school", "back to school letter", "back to school"],
  "extracted_fields": {
    "principal_name": "Dr. Johnson"
  },
  "confidence": 0.95,
  "summary": "User wants to update a back-to-school letter, changing the principal name to Dr. Johnson"
}

User: "I want to see some of my most recent letters"
{
  "intent": "find_document",
  "document_type": "letter",
  "search_terms": ["letter", "correspondence", "notice"],
  "extracted_fields": {},
  "confidence": 0.6,
  "summary": "User wants to browse their recent letters"
}

User: "Show me what documents I have"
{
  "intent": "find_document",
  "document_type": "unknown",
  "search_terms": ["memo", "letter", "notice", "report"],
  "extracted_fields": {},
  "confidence": 0.4,
  "summary": "User wants to browse their documents broadly"
}"""


# =============================================================================
# v6 SEARCH INTENT EXTRACTION (simplified — search only)
# =============================================================================

def extract_search_intent(
    user_prompt: str,
    azure_endpoint: Optional[str] = None,
    azure_api_key: Optional[str] = None,
    azure_deployment: Optional[str] = None,
    api_version: Optional[str] = None
) -> Dict[str, Any]:
    """
    Extract search intent only — no field extraction.
    Simplified for the v6 pipeline where search is separated from editing.

    Args:
        user_prompt: The user's natural language request
        azure_endpoint: Azure OpenAI endpoint (or env AZURE_OPENAI_ENDPOINT)
        azure_api_key: API key (or env AZURE_OPENAI_KEY)
        azure_deployment: Deployment name (or env AZURE_OPENAI_DEPLOYMENT)
        api_version: API version (or env AZURE_OPENAI_API_VERSION)

    Returns:
        Dictionary with:
        - document_type: str
        - search_terms: List[str]
        - summary: str
        - confidence: float
    """
    endpoint = azure_endpoint or os.environ.get('AZURE_OPENAI_ENDPOINT')
    api_key = azure_api_key or os.environ.get('AZURE_OPENAI_KEY')
    deployment = azure_deployment or os.environ.get('AZURE_OPENAI_DEPLOYMENT', 'gpt-4o-mini')
    version = api_version or os.environ.get('AZURE_OPENAI_API_VERSION', '2025-01-01-preview')

    if not endpoint or not api_key:
        raise ValueError(
            "Azure OpenAI credentials not configured. "
            "Set AZURE_OPENAI_ENDPOINT and AZURE_OPENAI_KEY environment variables."
        )

    client = AzureOpenAI(
        azure_endpoint=endpoint,
        api_key=api_key,
        api_version=version
    )

    try:
        response = client.chat.completions.create(
            model=deployment,
            messages=[
                {"role": "system", "content": SEARCH_INTENT_PROMPT},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.1,
            max_tokens=500,
            response_format={"type": "json_object"}
        )

        result_text = response.choices[0].message.content
        result = json.loads(result_text)

        # Hard cap: Graph API search(q='...') works best with short queries.
        # Too many terms joined with spaces returns zero results because
        # the API tries to match all terms together.
        search_terms = result.get("search_terms", [])[:3]
        context_search_terms = result.get("context_search_terms", [])[:4]

        logging.info(f"Search terms (capped): {search_terms}")
        logging.info(f"Context search terms (capped): {context_search_terms}")

        return {
            "document_type": result.get("document_type", "unknown"),
            "search_terms": search_terms,
            "context_search_terms": context_search_terms,
            "summary": result.get("summary", ""),
            "confidence": result.get("confidence", 0.5)
        }

    except Exception as e:
        logging.error(f"Error extracting search intent: {e}")
        raise


# =============================================================================
# v5 INTENT EXTRACTION (kept for backward compatibility / single-doc fallback)
# =============================================================================

def extract_intent(
    user_prompt: str,
    azure_endpoint: Optional[str] = None,
    azure_api_key: Optional[str] = None,
    azure_deployment: Optional[str] = None,
    api_version: Optional[str] = None
) -> Dict[str, Any]:
    """
    Extract intent and field values from a user prompt.
    This is the v5 function, kept for backward compatibility and the
    single-document fallback path.

    Args:
        user_prompt: The user's natural language request
        azure_endpoint: Azure OpenAI endpoint (or env AZURE_OPENAI_ENDPOINT)
        azure_api_key: Azure OpenAI API key (or env AZURE_OPENAI_KEY)
        azure_deployment: Deployment name (or env AZURE_OPENAI_DEPLOYMENT)
        api_version: API version (or env AZURE_OPENAI_API_VERSION)

    Returns:
        Dictionary with:
        - intent: str
        - document_type: str
        - search_terms: List[str]
        - extracted_fields: Dict[str, str]
        - confidence: float
        - summary: str
    """
    endpoint = azure_endpoint or os.environ.get('AZURE_OPENAI_ENDPOINT')
    api_key = azure_api_key or os.environ.get('AZURE_OPENAI_KEY')
    deployment = azure_deployment or os.environ.get('AZURE_OPENAI_DEPLOYMENT', 'gpt-4o-mini')
    version = api_version or os.environ.get('AZURE_OPENAI_API_VERSION', '2025-01-01-preview')

    if not endpoint or not api_key:
        raise ValueError(
            "Azure OpenAI credentials not configured. "
            "Set AZURE_OPENAI_ENDPOINT and AZURE_OPENAI_KEY environment variables."
        )

    client = AzureOpenAI(
        azure_endpoint=endpoint,
        api_key=api_key,
        api_version=version
    )

    try:
        response = client.chat.completions.create(
            model=deployment,
            messages=[
                {"role": "system", "content": INTENT_EXTRACTION_PROMPT},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.1,
            max_tokens=1000,
            response_format={"type": "json_object"}
        )

        result_text = response.choices[0].message.content
        result = json.loads(result_text)

        return {
            "intent": result.get("intent", "unknown"),
            "document_type": result.get("document_type", "unknown"),
            "search_terms": result.get("search_terms", []),
            "extracted_fields": result.get("extracted_fields", {}),
            "confidence": result.get("confidence", 0.5),
            "summary": result.get("summary", "")
        }

    except Exception as e:
        logging.error(f"Error extracting intent: {e}")
        raise


# =============================================================================
# FIELD MERGING (kept from v5)
# =============================================================================

def merge_fields(
    original_fields: List[Dict[str, Any]],
    extracted_fields: Dict[str, str]
) -> List[Dict[str, Any]]:
    """
    Merge AI-extracted field values into the original document fields.

    Args:
        original_fields: Fields extracted from document analysis
            Each field has: field_name, field_label, current_value, field_type, required
        extracted_fields: Field values extracted from user prompt
            Maps field_name to new value

    Returns:
        Updated fields list with new_value populated where matches found
    """
    merged = []

    for field in original_fields:
        field_copy = field.copy()
        field_name = field.get('field_name', '')

        if field_name in extracted_fields:
            field_copy['new_value'] = extracted_fields[field_name]
            field_copy['pre_filled'] = True
        else:
            field_copy['new_value'] = field.get('current_value', '')
            field_copy['pre_filled'] = False

        merged.append(field_copy)

    return merged


# =============================================================================
# ADAPTIVE CARD GENERATION (kept from v5)
# =============================================================================

def generate_field_input_card(fields: List[Dict[str, Any]], document_type: str) -> Dict:
    """
    Generate an Adaptive Card for editing document fields.
    Pre-fills values that were extracted from the user's prompt.

    Args:
        fields: List of fields with current_value and optional new_value
        document_type: Type of document (for display)

    Returns:
        Adaptive Card JSON as dictionary
    """
    body = [
        {
            "type": "TextBlock",
            "text": f"Update {document_type.replace('_', ' ').title()}",
            "weight": "Bolder",
            "size": "Large",
            "wrap": True
        },
        {
            "type": "TextBlock",
            "text": "Review and edit the fields below. Pre-filled values are from your request.",
            "wrap": True,
            "isSubtle": True,
            "spacing": "Small"
        }
    ]

    for field in fields:
        field_name = field.get('field_name', 'field')
        field_label = field.get('field_label', field_name)
        field_type = field.get('field_type', 'text')
        current_value = field.get('current_value', '')
        new_value = field.get('new_value', current_value)
        pre_filled = field.get('pre_filled', False)

        label_text = f"{field_label}"
        if pre_filled:
            label_text += " (pre-filled from your request)"

        body.append({
            "type": "TextBlock",
            "text": label_text,
            "weight": "Bolder",
            "spacing": "Medium"
        })

        if field_type == 'date':
            body.append({
                "type": "Input.Date",
                "id": field_name,
                "value": new_value if new_value else None
            })
        elif field_type == 'multiline':
            body.append({
                "type": "Input.Text",
                "id": field_name,
                "isMultiline": True,
                "value": new_value,
                "placeholder": f"Current: {current_value[:50]}..." if len(current_value) > 50 else f"Current: {current_value}"
            })
        else:
            body.append({
                "type": "Input.Text",
                "id": field_name,
                "value": new_value,
                "placeholder": f"Current: {current_value}"
            })

    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.5",
        "body": body,
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Generate Updated Document",
                "style": "positive",
                "data": {"action": "generate"}
            },
            {
                "type": "Action.Submit",
                "title": "Cancel",
                "data": {"action": "cancel"}
            }
        ]
    }

    return card

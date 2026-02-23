# Refresh v6 — M365 Agents Toolkit Deployment Guide

**Version:** 6.0 (M365AT)
**Last updated:** February 22, 2026
**Author:** Cameron Fadjo + Claude

---

## Overview

Refresh is an AI-powered document research and generation agent for K-12 school districts. It analyzes previous versions of recurring documents (back-to-school letters, budget memos, policy updates), discovers recent organizational changes, and generates updated versions grounded in document history and context.

This version replaces the Copilot Studio + Power Automate architecture with a **declarative agent** for the M365 Agents Toolkit, enabling deployment to the **Microsoft Agent Store**.

### Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    Microsoft 365 Copilot                     │
│                                                              │
│  User ──→ Declarative Agent ──→ API Plugin                  │
│       (declarativeAgent.json)    (openapi.yaml)              │
└──────────────────────────┬──────────────────────────────────┘
                           │ OAuth 2.0 Bearer Token
                           ▼
┌─────────────────────────────────────────────────────────────┐
│                  Azure Functions (Python)                     │
│                                                              │
│  ┌──────────────────┐  ┌──────────────────┐                 │
│  │ search-onedrive  │  │ extract-search-  │                 │
│  │ (Graph API)      │  │ intent (OpenAI)  │                 │
│  └──────────────────┘  └──────────────────┘                 │
│  ┌──────────────────┐  ┌──────────────────┐                 │
│  │ retrieve-and-    │  │ generate-from-   │                 │
│  │ analyze (Graph + │  │ synthesis        │                 │
│  │ OpenAI)          │  │ (OpenAI + Blob)  │                 │
│  └──────────────────┘  └──────────────────┘                 │
│  ┌──────────────────┐                                       │
│  │ save-to-onedrive │   graph_client.py (OBO token exchange)│
│  │ (Graph API)      │                                       │
│  └──────────────────┘                                       │
└──────────────────────────┬──────────────────────────────────┘
                           │
              ┌────────────┼────────────┐
              ▼            ▼            ▼
        ┌──────────┐ ┌──────────┐ ┌──────────┐
        │ Microsoft│ │  Azure   │ │  Azure   │
        │ Graph API│ │  OpenAI  │ │   Blob   │
        │ (OneDrive│ │ (GPT-4o  │ │ Storage  │
        │  files)  │ │  mini)   │ │          │
        └──────────┘ └──────────┘ └──────────┘
```

### What Changed from Copilot Studio

| Before (Copilot Studio) | After (M365 Agents Toolkit) |
|---|---|
| 10-node Copilot Studio topic | `declarativeAgent.json` instructions (declarative) |
| Flow 1: HTTP to extract-search-intent | `extractSearchIntent` plugin function (same endpoint) |
| Flow 2: Graph API search via PA connector | `searchOneDrive` plugin function (new endpoint) |
| Flow 3: Retrieve docs + analyze via PA | `retrieveAndAnalyze` plugin function (new endpoint) |
| Flow 4: Generate via HTTP | `generateFromSynthesis` plugin function (same endpoint) |
| Flow 5: Save to OneDrive via PA connector | `saveToOneDrive` plugin function (new endpoint) |
| Power Automate Graph connector auth | OAuth 2.0 OBO token exchange in Python |

### What Stayed the Same

All 5 core Python modules carry forward unchanged — 2,040+ lines of battle-tested code:

- `document_analyzer.py` — Text extraction + field analysis
- `document_generator.py` — Word document creation + synthesis generation
- `family_analyzer.py` — Cross-document comparative analysis engine
- `intent_extractor.py` — Azure OpenAI search intent extraction
- `blob_storage.py` — Azure Blob Storage + SAS URL generation

---

## Prerequisites

Before starting, ensure you have:

- **Microsoft 365 developer tenant** with M365 Copilot license
- **Azure subscription** with permissions to create resources
- **Azure OpenAI resource** with `gpt-4o-mini` model deployed
- **Azure Storage account** with a blob container named `generated-documents`
- **VS Code** with the **Microsoft 365 Agents Toolkit** extension installed
- **Node.js 18+** (for Agents Toolkit CLI)
- **Python 3.11** (for Azure Functions runtime)
- **Azure Functions Core Tools v4** (`npm install -g azure-functions-core-tools@4`)

---

## File Structure

```
Refresh-M365AT/
├── appPackage/                        # Agent manifests (packaged by toolkit)
│   ├── manifest.json                  # Teams app manifest (v1.19)
│   ├── declarativeAgent.json          # Agent instructions + conversation starters
│   ├── ai-plugin.json                 # API plugin manifest (v2.2, runtimes pattern)
│   ├── color.png                      # App icon — full color (192×192)
│   ├── outline.png                    # App icon — outline only (32×32)
│   ├── apiSpecificationFile/
│   │   └── openapi.yaml               # OpenAPI 3.0 spec for all endpoints
│   └── adaptiveCards/                 # Rich response cards for each operation
│       ├── extractSearchIntent.json
│       ├── searchOneDrive.json
│       ├── retrieveAndAnalyze.json
│       ├── generateFromSynthesis.json
│       └── saveToOneDrive.json
├── azure_function/                    # Python backend (deploy to Azure)
│   ├── function_app.py                # 13 API endpoints (3 new for M365AT)
│   ├── graph_client.py                # NEW — Graph API + OBO auth helper
│   ├── document_analyzer.py           # Text extraction + field analysis
│   ├── document_generator.py          # Word document creation
│   ├── family_analyzer.py             # Cross-document comparative analysis
│   ├── intent_extractor.py            # Azure OpenAI intent extraction
│   ├── blob_storage.py                # Azure Blob Storage + SAS URLs
│   ├── requirements.txt               # Python dependencies (includes msal)
│   ├── host.json                      # Functions runtime config
│   ├── local.settings.json            # Local dev environment variables
│   └── .funcignore                    # Deployment exclusions
├── env/                               # Environment configs (auto-populated)
│   ├── .env.dev                       # Dev environment variables
│   └── .env.local                     # Local development variables
├── m365agents.yml                     # Agents Toolkit project config (v1.10)
└── DEPLOYMENT_GUIDE.md                # This file
```

---

## Part 1: Open the Project in VS Code

1. Open VS Code
2. **File → Open Folder** → select the `Refresh-M365AT` folder
3. Install the **Microsoft 365 Agents Toolkit** extension if not already installed (search "Microsoft 365 Agents Toolkit" in the Extensions marketplace)
4. You should see the Agents Toolkit icon in the left sidebar

> At this point, don't run any Agents Toolkit commands yet — we need to set up the Azure backend and Entra ID first.

---

## Part 2: Entra ID App Registration (OAuth 2.0 OBO)

The declarative agent needs OAuth 2.0 On-Behalf-Of (OBO) flow to access the user's OneDrive. This requires an Entra ID app registration.

### 2.1 Create App Registration

1. Go to **Azure Portal** → **Microsoft Entra ID** → **App registrations**
2. Click **New registration**
3. Name: `Refresh v6 Agent`
4. Supported account types: **Accounts in this organizational directory only**
5. Redirect URI: leave blank (not needed for the OBO flow — the M365 Agents Toolkit handles this)
6. Click **Register**
7. **Copy the Application (client) ID** — you'll need this as `ENTRA_CLIENT_ID`
8. **Copy the Directory (tenant) ID** — you'll need this as `ENTRA_TENANT_ID`

### 2.2 Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission** → **Microsoft Graph** → **Delegated permissions**
3. Add these permissions:
   - `Files.ReadWrite` — read and write user's OneDrive files
   - `User.Read` — read user's basic profile
4. Click **Grant admin consent for [your tenant]**
5. Verify both permissions show a green checkmark under "Status"

### 2.3 Expose an API (Enable OBO Flow)

This is the critical step that allows the Azure Function to exchange the user's SSO token for a Graph API token.

1. Go to **Expose an API**
2. Click **Set** next to Application ID URI
3. Set it to: `api://YOUR-FUNCTION-APP-DOMAIN/YOUR-CLIENT-ID`
   - Example: `api://refresh-m365at.azurewebsites.net/a1b2c3d4-e5f6-7890-abcd-ef1234567890`
   - You'll update the domain once you create the Function App in Part 3
4. Click **Add a scope**:
   - Scope name: `access_as_user`
   - Who can consent: **Admins and users**
   - Admin consent display name: `Access Refresh API as user`
   - Admin consent description: `Allows M365 Copilot to call the Refresh API on behalf of the signed-in user`
   - State: **Enabled**
5. Click **Add scope**

### 2.4 Authorize Client Applications

This step is critical — it tells Entra ID which Microsoft apps are allowed to request tokens on behalf of the user. Without this, M365 Copilot and Teams won't be able to call your plugin.

1. Still on **Expose an API**, scroll down to **Authorized client applications**
2. Click **Add a client application** and add each of the following, checking the `access_as_user` scope for each:

| Client ID | Application |
|---|---|
| `1fec8e78-bce4-4aaf-ab1b-5451cc387264` | Teams Desktop & Mobile |
| `5e3ce6c0-2b1f-4285-8d4b-75ee78787346` | Teams Web |
| `4765445b-32c6-49b0-83e6-1115210e106b` | Microsoft 365 (Office Web) |
| `0ec893e0-5785-4de6-99da-4ed124e5296c` | Microsoft 365 (Office Desktop) |
| `d3590ed6-52b3-4102-aedd-aad2292ab01c` | Outlook Desktop |
| `bc59ab01-8403-45c6-8796-ac3ef710b3e3` | Outlook Web |
| `27922004-5251-4030-b22d-91ecd9a37ea4` | Outlook Mobile |

3. After adding all client applications, each should appear in the list with the `access_as_user` scope checked

> **Why this matters:** When M365 Copilot calls your plugin, it sends an SSO token. Your Azure Function exchanges that token via OBO flow. Entra ID will reject the exchange unless the originating client app (e.g., Teams) is pre-authorized here.

### 2.5 Create Client Secret

1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Description: `Refresh v6 Agent`
4. Expiry: 24 months (or your preference)
5. Click **Add**
6. **Copy the secret Value immediately** — it only shows once. This is your `ENTRA_CLIENT_SECRET`

### 2.6 Values to Save

You should now have these three values:

| Value | Where to find it | Used as |
|---|---|---|
| Application (client) ID | App registration → Overview | `ENTRA_CLIENT_ID` |
| Directory (tenant) ID | App registration → Overview | `ENTRA_TENANT_ID` |
| Client secret value | Certificates & secrets | `ENTRA_CLIENT_SECRET` |

---

## Part 3: Azure Function App Setup

### 3.1 Create the Function App

1. Go to **Azure Portal** → **Function App** → **Create**
2. Configure:
   - **Subscription:** Your Azure subscription
   - **Resource Group:** Create new or use existing
   - **Function App name:** `refresh-m365at` (or your preferred name)
   - **Runtime stack:** Python
   - **Version:** 3.11
   - **Region:** Same as your Azure OpenAI resource
   - **Operating System:** Linux
   - **Plan type:** Consumption (Serverless)
3. Click **Review + create** → **Create**
4. Wait for deployment to complete

### 3.2 Configure Environment Variables

Go to your Function App → **Settings** → **Environment variables** (or Configuration → Application settings) and add ALL of the following:

| Variable | Value | Purpose |
|---|---|---|
| `AzureWebJobsFeatureFlags` | `EnableWorkerIndexing` | **CRITICAL** — Required for Python v2 functions to register |
| `AZURE_OPENAI_ENDPOINT` | `https://YOUR-RESOURCE.openai.azure.com/` | Azure OpenAI endpoint URL |
| `AZURE_OPENAI_KEY` | Your API key | Azure OpenAI authentication |
| `AZURE_OPENAI_DEPLOYMENT` | `gpt-4o-mini` | Model deployment name |
| `AZURE_OPENAI_DEPLOYMENT_LARGE` | (optional, e.g. `gpt-4o`) | For complex analysis tasks |
| `AZURE_OPENAI_API_VERSION` | `2025-01-01-preview` | API version |
| `AZURE_STORAGE_CONNECTION_STRING` | Your storage connection string | Blob storage for generated docs |
| `AZURE_STORAGE_CONTAINER_NAME` | `generated-documents` | Container name |
| `SAS_TOKEN_EXPIRY_HOURS` | `24` | SAS URL lifetime |
| `ENTRA_TENANT_ID` | From Part 2.6 | OAuth tenant |
| `ENTRA_CLIENT_ID` | From Part 2.6 | OAuth client ID |
| `ENTRA_CLIENT_SECRET` | From Part 2.6 | OAuth client secret |

Click **Save** (this will restart the Function App).

> **IMPORTANT:** If you skipped `AzureWebJobsFeatureFlags` = `EnableWorkerIndexing`, your functions will deploy but never register. This is the most common deployment issue.

### 3.3 Update Entra ID Application ID URI

Now that you have the Function App domain, go back to Entra ID:

1. **Entra ID** → **App registrations** → **Refresh v6 Agent** → **Expose an API**
2. Update the Application ID URI to: `api://refresh-m365at.azurewebsites.net/YOUR-CLIENT-ID`
   - Replace `refresh-m365at` with your actual Function App name

### 3.4 Deploy Azure Functions

**Option A: Deploy from VS Code**

1. In VS Code, open the `azure_function` subfolder (**not** the project root — `host.json` must be at the root of what gets deployed)
2. Press **Ctrl+Shift+P** → search **Azure Functions: Deploy to Function App**
3. Select your `refresh-m365at` Function App
4. Confirm the deployment

**Option B: Deploy from CLI**

```bash
# Make sure you're logged in first
az login

# Deploy from INSIDE the azure_function folder
cd Refresh-M365AT/azure_function
func azure functionapp publish refresh-m365at
```

> **Common mistake:** If you deploy from the project root instead of `azure_function/`, you'll get "Cannot find required host.json file." Always deploy from inside `azure_function/`.

### 3.5 Verify Deployment

After deployment, verify the health endpoint. **Do not proceed to Part 5 until this returns a healthy response.**

```
GET https://refresh-m365at.azurewebsites.net/api/health
```

Expected response:
```json
{
    "status": "healthy",
    "version": "6.0",
    "services": {
        "azure_openai": true,
        "azure_openai_large_model": false,
        "blob_storage": true
    },
    "endpoints": [
        "POST /api/search-onedrive",
        "POST /api/retrieve-and-analyze",
        "POST /api/save-to-onedrive",
        "POST /api/extract-search-intent",
        ...
    ]
}
```

If the health endpoint returns the full list of 13 endpoints, your deployment is successful.

**If health returns 404 or functions aren't listed:**
1. Check Azure Portal → Function App → **Functions** — all 13 should appear
2. If none appear, verify `AzureWebJobsFeatureFlags` = `EnableWorkerIndexing` in app settings
3. Check **Deployment center** → **Logs** for errors
4. Try redeploying with: `func azure functionapp publish refresh-m365at --build remote`

### All Endpoints

> **Note:** The "v5 compat" endpoints are for backwards compatibility with the older Copilot Studio architecture. They are not used by the M365AT agent and can be ignored.

| Group | Endpoint | Method | Purpose |
|---|---|---|---|
| **M365AT** | `/api/search-onedrive` | POST | Search user's OneDrive via Graph API |
| **M365AT** | `/api/retrieve-and-analyze` | POST | Fetch docs + comparative analysis |
| **M365AT** | `/api/save-to-onedrive` | POST | Save generated doc to OneDrive |
| **v6** | `/api/extract-search-intent` | POST | Parse user request into search terms |
| **v6** | `/api/analyze-family` | POST | Cross-document comparative analysis |
| **v6** | `/api/generate-from-synthesis` | POST | Generate document from analysis |
| **v5 compat** | `/api/extract-intent` | POST | Intent extraction (with field extraction) |
| **v5 compat** | `/api/analyze-document` | POST | Single-document analysis |
| **v5 compat** | `/api/generate-document` | POST | Generate document from fields |
| **v5 compat** | `/api/merge-fields` | POST | Merge user changes into fields |
| **v5 compat** | `/api/refresh-document` | POST | Combined analyze + generate |
| **Utility** | `/api/health` | GET | Health check |
| **Utility** | `/api/storage-status` | GET | Blob storage status |

---

## Part 4: Configure the Agent Package

Now that your Azure Function App is deployed and Entra ID is configured, update the agent package files with your actual values.

### 4.1 Replace Placeholders

The agent package files contain `${{...}}` placeholders. Some you replace manually; others the Agents Toolkit fills in automatically during provisioning.

**Replace these NOW (before provisioning):**

| Placeholder | Replace with | File |
|---|---|---|
| `${{AZURE_FUNCTION_URL}}` | `https://refresh-m365at.azurewebsites.net` (your Function App URL, **no** trailing `/api`) | `appPackage/apiSpecificationFile/openapi.yaml` |
| `${{AZURE_FUNCTION_DOMAIN}}` | `refresh-m365at.azurewebsites.net` (domain only, no `https://`) | `appPackage/manifest.json` |
| `${{ENTRA_TENANT_ID}}` | Your tenant ID GUID (from Part 2.6) | `appPackage/apiSpecificationFile/openapi.yaml` |
| `${{ENTRA_CLIENT_ID}}` | Your client ID GUID (from Part 2.6) | `appPackage/manifest.json` |

**DO NOT replace these — the toolkit fills them in during provisioning:**

| Placeholder | Populated by | File |
|---|---|---|
| `${{TEAMS_APP_ID}}` | `teamsApp/create` step | `appPackage/manifest.json` |
| `${{OAUTH2_REGISTRATION_ID}}` | `oauth/register` step | `appPackage/ai-plugin.json` |
| `${{APP_NAME_SUFFIX}}` | Environment config (blank in prod, `-dev` in dev) | `appPackage/manifest.json` |
| `${{TEAMSFX_ENV}}` | Toolkit runtime | `m365agents.yml` |
| `${{AGENT_SCOPE}}` | Toolkit runtime | `m365agents.yml` |

> **Important:** Leave the auto-populated placeholders exactly as `${{...}}`. If you replace them manually, provisioning may fail or overwrite your values.

### 4.2 App Icons

The agent requires two icon files in `appPackage/`:

| File | Size | Description |
|---|---|---|
| `color.png` | 192×192 px | Full-color icon (shown in Teams app gallery, Copilot agent list) |
| `outline.png` | 32×32 px | Transparent background, white foreground only (shown in Teams message bar) |

Placeholder icons are included in the project. Replace them with your organization's branding before publishing to the Agent Store.

### 4.3 Customize the Agent (Optional)

**Conversation starters** — In `appPackage/declarativeAgent.json`, edit the `conversation_starters` array to match the document types your district uses:

```json
"conversation_starters": [
    {"title": "Back-to-School Letter", "text": "I need this year's back-to-school letter"},
    {"title": "Budget Memo", "text": "Update our budget memo for the new fiscal year"},
    ...
]
```

**Agent instructions** — The `instructions` field in `appPackage/declarativeAgent.json` contains the 6-step workflow the agent follows. You can customize the language, add district-specific guidance, or adjust the workflow. Max 8,000 characters.

**Branding** — In `appPackage/manifest.json`, update:
- `developer.name` — your organization name
- `developer.websiteUrl` — your organization URL
- `accentColor` — your brand color hex code

**Adaptive cards** — Each API response renders as a rich card in Teams/Copilot. The card templates are in `appPackage/adaptiveCards/`. Customize their layout, fields, and styling as needed.

---

## Part 5: Build and Deploy the Agent

### 5.1 Provision

In VS Code with the Agents Toolkit extension:

1. Click the **Agents Toolkit** icon in the left sidebar
2. Under **Lifecycle**, click **Provision**
3. Select the **dev** environment (if not visible, ensure `env/.env.dev` exists)
4. Sign in to your M365 tenant when prompted
5. The toolkit will execute the steps in `m365agents.yml`:
   - Register OAuth connection (`oauth/register`)
   - Create/update the Teams app
   - Generate `TEAMS_APP_ID` and `OAUTH2_REGISTRATION_ID`
   - Package the manifests from `appPackage/`
   - Extend the app to M365

### 5.2 Sideload for Testing

1. In the Agents Toolkit sidebar, click **Run and Debug** → **Preview in Copilot (Edge)** or **Preview in Teams**
2. Or manually sideload:
   - Go to **Teams** → **Apps** → **Manage your apps** → **Upload a custom app**
   - Select the `.zip` from `build/appPackage/`

### 5.3 Test the Full Workflow

1. Open **M365 Copilot** in Teams (or Outlook)
2. Find and activate the **Refresh Agent**
3. Try: `"I need this year's back-to-school letter"`
4. The agent should:
   - Call `extractSearchIntent` → identify document type + search terms
   - Call `searchOneDrive` twice → find document family + context docs
   - Call `retrieveAndAnalyze` → download files, run comparative analysis
   - Present analysis summary → stable elements, variable elements, org context
   - Ask for confirmation/changes
   - Call `generateFromSynthesis` → create new document
   - Call `saveToOneDrive` → save to OneDrive
   - Present the OneDrive link

### 5.4 Publish to Agent Store

When ready to publish:

1. In the Agents Toolkit sidebar, under **Lifecycle**, click **Publish**
2. Go to **Teams Admin Center** → **Manage apps** → approve the submission
3. For the Microsoft commercial marketplace, your agent must pass:
   - Responsible AI (RAI) validation
   - Microsoft Marketplace certification
   - Teams validation guidelines

---

## Part 6: API Reference

### Authentication Requirements

Not all endpoints require the OBO token. The agent plugin sends the Bearer token automatically for all calls, but for debugging purposes:

| Endpoint | Requires OBO Token? | Why |
|---|---|---|
| `extract-search-intent` | No | Calls Azure OpenAI only |
| `search-onedrive` | **Yes** | Calls Microsoft Graph API |
| `retrieve-and-analyze` | **Yes** | Calls Microsoft Graph API |
| `generate-from-synthesis` | No | Calls Azure OpenAI + Blob Storage |
| `save-to-onedrive` | **Yes** | Calls Microsoft Graph API |

If you see 401/403 errors, check whether the failing endpoint is one that requires OBO. If so, the issue is in your Entra ID configuration (Part 2). If not, the issue is likely Azure OpenAI or Blob Storage credentials (Part 3.2).

### `POST /api/extract-search-intent`

Parses a user's natural language request into structured search terms.

**Request:**
```json
{
    "prompt": "I need this year's back-to-school letter"
}
```

**Response:**
```json
{
    "success": true,
    "document_type": "back_to_school_letter",
    "search_terms": ["back to school", "letter", "welcome"],
    "context_search_terms": ["new staff 2026", "budget update", "technology initiative"],
    "summary": "User needs a back-to-school letter for this year",
    "confidence": 0.9
}
```

### `POST /api/search-onedrive`

Searches the user's OneDrive via Microsoft Graph API.

**Headers:** `Authorization: Bearer <SSO token>`

**Request:**
```json
{
    "search_terms": "back to school letter"
}
```

**Response:**
```json
{
    "success": true,
    "documents": [
        {
            "id": "01ABC123DEF456",
            "name": "Back to School Letter 2025.docx",
            "path": "/drive/root:/Documents",
            "webUrl": "https://contoso-my.sharepoint.com/...",
            "lastModified": "2025-08-14T10:30:00Z",
            "createdDateTime": "2025-08-10T09:00:00Z",
            "size": 45632
        }
    ],
    "count": 3
}
```

### `POST /api/retrieve-and-analyze`

Fetches documents from OneDrive and runs cross-document comparative analysis.

**Headers:** `Authorization: Bearer <SSO token>`

**Request:**
```json
{
    "document_ids": ["01ABC123", "01DEF456", "01GHI789"],
    "context_document_ids": ["01JKL012", "01MNO345"],
    "user_context": "I need this year's back-to-school letter"
}
```

**Response:**
```json
{
    "success": true,
    "family_type": "back_to_school_letter",
    "family_type_display": "Back-to-School Letter",
    "document_count": 3,
    "date_range": "2023-2025",
    "analysis": {
        "stable_elements": { "...": "..." },
        "variable_elements": { "...": "..." },
        "emerging_elements": { "...": "..." }
    },
    "recommended_base": "Back to School Letter 2025.docx",
    "base_document_text": "Dear Families, Welcome to the 2025-2026 school year...",
    "organizational_context": "New assistant principal Dr. Johnson hired; 1:1 device program expanding to all grades",
    "confidence": 0.9,
    "summary": "Analyzed 3 back-to-school letters spanning 2023-2025"
}
```

### `POST /api/generate-from-synthesis`

Generates a new document version from comparative analysis.

**Request:**
```json
{
    "family_analysis": { "...analysis object from above..." },
    "base_document_text": "Dear Families, Welcome to the 2025-2026 school year...",
    "organizational_context": "New assistant principal Dr. Johnson hired...",
    "user_changes": "Change the principal name to Dr. Johnson",
    "target_year": "2026-2027"
}
```

**Response:**
```json
{
    "success": true,
    "generated_text": "Dear Families, Welcome to the 2026-2027 school year...",
    "changes_applied": [
        "Updated school year to 2026-2027",
        "Updated principal to Dr. Johnson",
        "Added 1:1 device program reference"
    ],
    "flags": [
        {"field": "first_day_of_school", "reason": "Date not found in source documents", "current_placeholder": "[First Day of School]"}
    ],
    "filename": "Back to School Letter - 2026-2027.docx",
    "download_url": "https://storageaccount.blob.core.windows.net/generated-documents/...?sv=...",
    "expires_in_hours": 24,
    "storage_type": "blob_sas_url"
}
```

### `POST /api/save-to-onedrive`

Saves a generated document to the user's OneDrive.

**Headers:** `Authorization: Bearer <SSO token>`

**Request:**
```json
{
    "download_url": "https://storageaccount.blob.core.windows.net/...?sv=...",
    "filename": "Back to School Letter - 2026-2027.docx",
    "folder_path": "Refresh"
}
```

**Response:**
```json
{
    "success": true,
    "savedPath": "/Refresh/Back to School Letter - 2026-2027.docx",
    "webUrl": "https://contoso-my.sharepoint.com/personal/...",
    "itemId": "01XYZ789"
}
```

---

## Part 7: Authentication Deep Dive

### How OBO Token Exchange Works

```
1. User signs into M365 (Teams/Outlook) via SSO
              ↓
2. M365 Copilot receives SSO token (JWT) with user's identity
              ↓
3. When agent calls a plugin endpoint, Copilot includes
   the SSO token in the Authorization header:
   Authorization: Bearer eyJ0eXAiOiJKV1Q...
              ↓
4. Azure Function extracts the Bearer token from the header
   (graph_client.extract_token_from_header)
              ↓
5. Azure Function exchanges it via MSAL OBO flow:
   - Sends to: https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
   - grant_type: urn:ietf:params:oauth:grant-type:jwt-bearer
   - assertion: {user's SSO token}
   - client_id: {ENTRA_CLIENT_ID}
   - client_secret: {ENTRA_CLIENT_SECRET}
   - scope: https://graph.microsoft.com/.default
   (graph_client.exchange_token)
              ↓
6. Azure AD returns a Graph API access token
              ↓
7. Azure Function calls Microsoft Graph API with the new token
   (graph_client.search_onedrive, get_file_content, etc.)
```

### Required Scopes

| Scope | Required by | Purpose |
|---|---|---|
| `Files.ReadWrite` | search-onedrive, retrieve-and-analyze, save-to-onedrive | Read/write OneDrive files |
| `User.Read` | General | Read user's basic profile |

### Troubleshooting Authentication

| Error | Cause | Fix |
|---|---|---|
| `AADSTS50013: Invalid assertion` | SSO token expired or malformed | Token should be fresh from M365 Copilot. If persistent, check that the Application ID URI matches |
| `AADSTS65001: User hasn't consented` | Admin consent not granted | Go to Entra ID → API permissions → Grant admin consent |
| `AADSTS700024: Client assertion contains an invalid signature` | Wrong client secret or client ID | Verify ENTRA_CLIENT_ID and ENTRA_CLIENT_SECRET in Function App settings |
| `401 Unauthorized` from Graph API | OBO token doesn't have required scopes | Check that Files.ReadWrite is in API permissions with admin consent |
| `403 Forbidden` from Graph API | Permission exists but admin consent missing | Grant admin consent in Entra ID |

---

## Part 8: The Agent Pipeline

### How the Declarative Agent Orchestrates

The `instructions` field in `declarativeAgent.json` tells M365 Copilot the exact workflow to follow. It replaces the 10-node Copilot Studio topic with a single set of natural language instructions.

**Step 1 — Understand the Request:** User asks for a document. Agent calls `extractSearchIntent` to parse the request into document type + search terms + context search terms.

**Step 2 — Search for Document Family:** Agent calls `searchOneDrive` with the search terms. Finds previous versions of the target document.

**Step 3 — Search for Organizational Context:** Agent calls `searchOneDrive` again with context search terms. Finds recent memos, announcements, and updates that may affect the document.

**Step 4 — Analyze:** Agent calls `retrieveAndAnalyze` with both sets of document IDs. The endpoint downloads all files from OneDrive, runs cross-document comparative analysis, and returns stable elements (things that never change), variable elements (things that change each year, with predicted new values), emerging elements (recently added sections), and organizational context extracted from context documents.

**Step 5 — Generate:** After user confirms and provides any changes, agent calls `generateFromSynthesis`. Returns the new document with a download URL and list of changes applied.

**Step 6 — Save:** Agent calls `saveToOneDrive` to save the generated document to the user's OneDrive Refresh folder.

### Complete Data Flow Example

```
User: "I need the back-to-school letter"

→ extractSearchIntent("I need the back-to-school letter")
  ← {document_type: "back_to_school_letter",
     search_terms: ["back to school", "letter"],
     context_search_terms: ["new staff 2026", "budget update"]}

→ searchOneDrive("back to school letter")
  ← {documents: [{id: "A1", name: "BTS 2025.docx"},
                  {id: "A2", name: "BTS 2024.docx"},
                  {id: "A3", name: "BTS 2023.docx"}], count: 3}

→ searchOneDrive("new staff 2026 budget update")
  ← {documents: [{id: "B1", name: "Staff Update Jan 2026.docx"},
                  {id: "B2", name: "Budget Memo Q2.docx"}], count: 2}

Agent: "I found 3 previous back-to-school letters and 2 recent
        organizational documents. Analyzing now..."

→ retrieveAndAnalyze({
    document_ids: ["A1", "A2", "A3"],
    context_document_ids: ["B1", "B2"],
    user_context: "I need the back-to-school letter"
  })
  ← {family_type: "back_to_school_letter",
     analysis: {stable, variable, emerging},
     base_document_text: "Dear Families...",
     organizational_context: "New AP Dr. Johnson; 1:1 device program"}

Agent: "Here's what I found across 3 versions (2023-2025):
        - Always includes: welcome message, principal greeting, school hours
        - Changes each year: dates, staff names, new programs
        - Recent additions: technology section (added 2024)
        - Org context: New AP Dr. Johnson hired, 1:1 devices expanding
        Shall I generate the 2026-2027 version?"

User: "Yes, and change the principal to Dr. Johnson"

→ generateFromSynthesis({
    family_analysis: {...},
    base_document_text: "Dear Families...",
    organizational_context: "New AP Dr. Johnson...",
    user_changes: "Change the principal to Dr. Johnson",
    target_year: "2026-2027"
  })
  ← {generated_text: "Dear Families, Welcome to 2026-2027...",
     changes_applied: ["Updated year", "Updated principal", "Added 1:1 devices"],
     flags: [{field: "first_day", reason: "Date unknown"}],
     filename: "Back to School Letter - 2026-2027.docx",
     download_url: "https://blob.../...?sv=..."}

Agent: "Generated! 3 changes applied, 1 field flagged for review.
        Save to OneDrive?"

User: "Yes"

→ saveToOneDrive({download_url: "...", filename: "Back to School Letter - 2026-2027.docx"})
  ← {savedPath: "/Refresh/Back to School Letter - 2026-2027.docx",
     webUrl: "https://contoso-my.sharepoint.com/..."}

Agent: "Saved! View your document: [link]"
```

---

## Part 9: Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| Functions not registering after deployment | Missing `AzureWebJobsFeatureFlags` | Add `EnableWorkerIndexing` to app settings |
| Health endpoint returns 404 | Deployment didn't complete | Check deployment logs, redeploy |
| 401 from search-onedrive | OBO token exchange failing | Check ENTRA_CLIENT_ID, ENTRA_CLIENT_SECRET, ENTRA_TENANT_ID |
| 403 from Graph API | Admin consent not granted | Grant admin consent in Entra ID |
| OneDrive search returns empty | Search terms too broad/long | `intent_extractor.py` caps at 3 terms; try simpler queries |
| retrieve-and-analyze times out | Too many/large documents | Limit to 5 documents; check Function App timeout settings |
| generate-from-synthesis returns 500 | Azure OpenAI config wrong | Verify AZURE_OPENAI_ENDPOINT, KEY, and DEPLOYMENT name |
| Agent doesn't appear in M365 Copilot | Not sideloaded or provisioned | Run Agents Toolkit → Provision → Preview in Teams |
| "Could not connect" in agent | Plugin OAuth not configured | Check `ai-plugin.json` runtimes auth config, `openapi.yaml` security scheme, and Entra ID setup |
| File not saving to OneDrive | Graph API write permission | Verify Files.ReadWrite scope with admin consent |
| Agent calls wrong endpoint | OpenAPI spec URL mismatch | Check server URL in `appPackage/apiSpecificationFile/openapi.yaml` matches your Function App |
| Import errors in Function App logs | pip dependencies not installed | Redeploy; if persists, try `func azure functionapp publish --build remote` |

---

## Appendix A: Environment Variables Quick Reference

| Variable | Example Value | Required |
|---|---|---|
| `AzureWebJobsFeatureFlags` | `EnableWorkerIndexing` | Yes |
| `AZURE_OPENAI_ENDPOINT` | `https://myoai.openai.azure.com/` | Yes |
| `AZURE_OPENAI_KEY` | `abc123...` | Yes |
| `AZURE_OPENAI_DEPLOYMENT` | `gpt-4o-mini` | Yes |
| `AZURE_OPENAI_DEPLOYMENT_LARGE` | `gpt-4o` | No (recommended) |
| `AZURE_OPENAI_API_VERSION` | `2025-01-01-preview` | Yes |
| `AZURE_STORAGE_CONNECTION_STRING` | `DefaultEndpointsProtocol=https;...` | Yes |
| `AZURE_STORAGE_CONTAINER_NAME` | `generated-documents` | Yes |
| `SAS_TOKEN_EXPIRY_HOURS` | `24` | Yes |
| `ENTRA_TENANT_ID` | `contoso.onmicrosoft.com` or GUID | Yes |
| `ENTRA_CLIENT_ID` | `a1b2c3d4-e5f6-...` | Yes |
| `ENTRA_CLIENT_SECRET` | `secret-value` | Yes |

---

## Appendix B: Key Agentic Behaviors

These are the intelligent behaviors built into the Refresh system — they carry forward from v6 unchanged:

1. **Two-phase search strategy** — Searches for both the target document family AND recent organizational documents that might affect it. The LLM decides what context to search for based on the document type.

2. **Chronological document reasoning** — Analyzes documents in date order to identify trends, not just current state. Understands that a field changed from X → Y → Z across three years.

3. **Stable/variable/emerging classification** — Categorizes every document element: stable (never changes), variable (changes predictably), emerging (recently added). This drives the generation strategy.

4. **Predictive value generation with flagging** — For variable elements, predicts the new value (e.g., next year's date) but flags low-confidence predictions with `[PLACEHOLDER]` for human review.

5. **Organizational context extraction** — Reads context documents (memos, announcements, budget updates) and synthesizes relevant changes into the generated document. Discovers things like "new principal hired" or "1:1 device program expanding."

6. **Base document selection** — Automatically selects the most recent, most complete version as the generation base. Doesn't blindly use the newest — considers completeness.

7. **Graceful degradation** — Works with 1 document (single-doc fallback), no context documents (skips org context), or even failed searches (suggests alternative terms). Never crashes on missing data.

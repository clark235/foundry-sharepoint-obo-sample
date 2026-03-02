# Pattern 2: M365 Copilot Retrieval API + Foundry Agent

This pattern separates **retrieval** (via the M365 Copilot Retrieval API) from **synthesis** (via an Azure AI Foundry agent). You control the full pipeline: acquire the user's token, search SharePoint through Graph, then feed the permission-trimmed results to Foundry for analysis.

## Why This Pattern?

- **Full orchestration control** — you decide what to retrieve, how to filter, and how to prompt the agent
- **Works without publishing to M365 Copilot** — runs as a standalone app
- **Mix data sources** — combine SharePoint results with other APIs before sending to Foundry
- **Audit trail** — you can log exactly what content was retrieved and what the agent saw

## How It Works

```
User signs in via MSAL (interactive browser)
    ↓
Delegated access token for Graph API
    ↓
POST to /beta/copilot/retrieval with user's token
    ↓
Graph returns permission-trimmed SharePoint chunks
(only content the user can access)
    ↓
Format chunks as context
    ↓
Send context + user query to Foundry agent
(agent uses service identity — DefaultAzureCredential)
    ↓
Agent synthesises answer
```

**Key insight:** The Retrieval API enforces permissions **before** content reaches your code. By the time you pass content to Foundry, it's already been trimmed to what the user can see. This means the Foundry agent doesn't need the user's identity — it can use a service credential.

## Prerequisites

1. **Azure AI Foundry project** with a pre-created agent (note the agent ID)
2. **M365 Copilot licence** on the user account — or pay-as-you-go Copilot metering enabled on the tenant
3. **Azure AD app registration** — public client with delegated permissions:
   - `Files.Read.All`
   - `Sites.Read.All`
4. **Python 3.9+**

## Setup

### 1. Create an Azure AD App Registration (Public Client)

1. Go to [Azure Portal → App registrations → New registration](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Name: `Foundry Retrieval Sample`
3. Supported account types: **Single tenant**
4. Redirect URI: `http://localhost` (type: Mobile and desktop applications)
5. Under **Authentication**, enable **Allow public client flows** → Yes
6. Under **API permissions**, add:
   - `Microsoft Graph → Delegated → Files.Read.All`
   - `Microsoft Graph → Delegated → Sites.Read.All`
7. Grant admin consent (or have users consent individually)

### 2. Create a Foundry Agent

Create an agent in your Foundry project (via portal or SDK) and note its ID (format: `asst_xxxxxxxxxxxxxxxxxxxx`).

### 3. Configure Environment

```bash
cp .env.example .env
# Edit .env with your values
```

### 4. Run

```bash
pip install -r requirements.txt
python main.py
```

## The Trust Model

```
┌─────────────┐     delegated token     ┌─────────────────┐
│    User      │ ──────────────────────→ │  Graph API      │
│  (browser)   │                         │  Retrieval API  │
└─────────────┘                          └────────┬────────┘
                                                  │
                                    permission-trimmed content
                                                  │
                                                  ↓
                                         ┌────────────────┐
                                         │  Your App      │
                                         │  (formats      │
                                         │   context)     │
                                         └────────┬───────┘
                                                  │
                                         service identity
                                                  │
                                                  ↓
                                         ┌────────────────┐
                                         │  Foundry Agent  │
                                         │  (synthesises   │
                                         │   answer)       │
                                         └────────────────┘
```

- **Graph enforces permissions** — the Retrieval API only returns content the signed-in user can access
- **Your app sees trimmed content** — no risk of accidentally exposing restricted documents
- **Foundry sees what you give it** — the agent operates on pre-approved content only
- **Two identity boundaries** — user token for Graph, service token for Foundry

## Retrieval API Details

**Endpoint:** `POST https://graph.microsoft.com/beta/copilot/retrieval`

**Rate limits:**
- 200 requests per user per hour
- Responses are chunked (not full documents)

**Data sources:** `sharepoint`, `onedrive`, `teams`

**Optional filtering:**
```json
{
  "queryString": "your search query",
  "dataSource": "sharepoint",
  "maximumNumberOfResults": 10,
  "filterExpression": "path:https://contoso.sharepoint.com/sites/HR"
}
```

## Limitations

- **Beta API** — the Retrieval API is in beta and may change
- **Rate limited** — 200 req/user/hour may not be enough for heavy use
- **Chunked results** — you get text chunks, not full documents
- **No streaming** — the Retrieval API returns all results at once
- **Requires user interaction** — the user must sign in interactively (or you need a refresh token)

## Files

| File | Description |
|---|---|
| `main.py` | Complete working example |
| `.env.example` | Environment variable template |
| `requirements.txt` | Python dependencies |

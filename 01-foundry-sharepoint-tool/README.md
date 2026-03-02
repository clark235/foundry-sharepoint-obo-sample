# Pattern 1: Foundry Agent with Native SharePoint Tool

This pattern uses Azure AI Foundry's built-in `SharepointToolDefinition` to give an agent access to SharePoint documents. The SDK handles **identity passthrough (OBO)** automatically — the user's permissions are enforced by SharePoint, so the agent can only see documents the user has access to.

## How It Works

```
User signs in (InteractiveBrowserCredential)
    ↓
AIProjectClient authenticates as the user
    ↓
Agent created with SharepointToolDefinition
    ↓
User sends a question
    ↓
Foundry passes user's OBO token to SharePoint
    ↓
SharePoint returns only documents the user can access
    ↓
Agent synthesises an answer with citations
```

**Key point:** The `SharepointToolDefinition` uses identity passthrough under the hood. You don't write any Graph API or SharePoint API calls — Foundry handles the token exchange, the search, and the grounding automatically.

## Prerequisites

1. **Azure AI Foundry project** — [Create one](https://ai.azure.com)
2. **M365 Copilot licence** on the user account — required for SharePoint grounding
3. **SharePoint connection** configured in Foundry (see setup below)
4. **Azure AD app registration** — for interactive user sign-in
5. **Python 3.9+**

## Setup

### 1. Create an Azure AD App Registration

1. Go to [Azure Portal → App registrations → New registration](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Name: `Foundry SharePoint Sample` (or similar)
3. Supported account types: **Single tenant**
4. Redirect URI: `http://localhost` (type: Mobile and desktop applications)
5. Note the **Application (client) ID** and **Directory (tenant) ID**

### 2. Create a SharePoint Connection in Foundry

1. Go to [Azure AI Foundry](https://ai.azure.com) → your project
2. **Project Settings** → **Connected resources** → **+ New connection**
3. Select **SharePoint**
4. Authenticate with an account that has SharePoint access
5. Note the **Connection ID** (format: `/subscriptions/.../connections/<name>`)

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

A browser window will open for sign-in. After authentication, the agent will query SharePoint and print results.

## How OBO Works Here

1. `InteractiveBrowserCredential` prompts the user to sign in via browser
2. The credential is passed to `AIProjectClient`, which uses it for all API calls
3. When the agent invokes the SharePoint tool, Foundry performs an OBO token exchange — it takes the user's token and requests a new token scoped to SharePoint
4. SharePoint receives a token that represents **the user**, not the app or Foundry service
5. SharePoint applies its normal permission checks and returns only accessible documents

This means: **if a user can't see a document in SharePoint, the agent can't see it either.**

## Limitations

- **Cannot be published to M365 Copilot** — when a Foundry agent is surfaced through M365 Copilot (e.g., in Teams), the identity chain breaks. The user's token doesn't flow through the M365 → Foundry boundary correctly. Use [Pattern 3](../03-declarative-agent-manifest/) for that scenario.
- **Interactive sign-in required** — this pattern requires the user to sign in via browser. For server-to-server scenarios, you'd need a different credential type with delegated permissions.
- **M365 Copilot licence required** — without it, the SharePoint tool won't return results even if the user has SharePoint access.
- **SharePoint connection scope** — the connection is project-level, so all agents in the project share it.

## Files

| File | Description |
|---|---|
| `main.py` | Complete working example |
| `.env.example` | Environment variable template |
| `requirements.txt` | Python dependencies |

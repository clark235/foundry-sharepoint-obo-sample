# Pattern 3: Declarative Agent Manifest + Foundry API Plugin

This pattern runs **inside M365 Copilot** (Teams, M365 Chat, etc.) as a declarative agent. SharePoint grounding is handled natively by M365 Copilot, and a Foundry-backed API plugin provides custom analysis capabilities. No Python needed for the agent itself — just JSON manifests.

## Why This Pattern?

- **Runs in M365 Copilot** — users access it directly in Teams or M365 Chat
- **SharePoint grounding is automatic** — M365 Copilot handles the OBO token exchange and permission trimming
- **No custom retrieval code** — you don't call the Retrieval API yourself
- **Extensible** — add Foundry-powered actions for custom logic beyond simple Q&A
- **Enterprise-ready** — deployed via Teams Developer Portal, manageable by IT admins

## How It Works

```
User asks a question in M365 Copilot
    ↓
Declarative agent instructions guide Copilot
    ↓
Copilot searches SharePoint (GraphConnectors capability)
    ↓
SharePoint returns permission-trimmed results (OBO handled by Copilot)
    ↓
If custom analysis needed: Copilot calls Foundry plugin endpoint
    ↓
Foundry endpoint receives query + SharePoint context
    ↓
Foundry agent processes and returns analysis
    ↓
Copilot presents the combined answer to the user
```

## Architecture

```
┌──────────────────────────────────────────────────┐
│  M365 Copilot (Teams / M365 Chat)                │
│                                                    │
│  ┌─────────────────────┐  ┌─────────────────────┐ │
│  │  SharePoint          │  │  Foundry Plugin     │ │
│  │  (GraphConnectors)   │  │  (API action)       │ │
│  │                      │  │                     │ │
│  │  OBO handled by      │  │  POST /analyze      │ │
│  │  M365 Copilot        │  │  → Azure App Svc    │ │
│  │  automatically        │  │  → Foundry Agent    │ │
│  └─────────────────────┘  └─────────────────────┘ │
└──────────────────────────────────────────────────┘
```

## Prerequisites

1. **M365 Copilot licence** on the user account
2. **Teams Developer Portal** access — to deploy the declarative agent
3. **Azure AI Foundry project** with a pre-created agent (for the plugin endpoint)
4. **Azure hosting** for the stub endpoint (App Service, Container App, or Functions)

## Files

| File | Description |
|---|---|
| `declarative-agent.json` | Declarative agent manifest — defines the agent's name, instructions, SharePoint capability, and Foundry action |
| `foundry-plugin.json` | OpenAPI spec for the Foundry analysis plugin — defines the `/analyze` endpoint |
| `stub-endpoint/app.py` | Flask app implementing the `/analyze` endpoint — calls a Foundry agent |
| `stub-endpoint/requirements.txt` | Python dependencies for the stub endpoint |

## Setup

### 1. Deploy the Stub Endpoint

The stub endpoint (`stub-endpoint/app.py`) needs to be hosted somewhere accessible by M365 Copilot.

**Option A: Azure App Service**
```bash
cd stub-endpoint
az webapp up --name your-foundry-endpoint --runtime PYTHON:3.11
az webapp config appsettings set --name your-foundry-endpoint \
  --settings FOUNDRY_PROJECT_ENDPOINT=https://... FOUNDRY_AGENT_ID=asst_...
```

**Option B: Azure Container Apps**
```bash
# Build and deploy a container
az containerapp up --name foundry-endpoint --source stub-endpoint/
```

**Option C: Azure Functions**
Adapt `app.py` to use Azure Functions HTTP trigger.

### 2. Update the Plugin URL

In `foundry-plugin.json`, replace the server URL:
```json
"servers": [
  { "url": "https://your-foundry-endpoint.azurewebsites.net" }
]
```

### 3. Deploy the Declarative Agent

**Via Teams Developer Portal:**
1. Go to [Teams Developer Portal](https://dev.teams.microsoft.com/)
2. Create a new app
3. Under **Copilot agents**, add a declarative agent
4. Upload `declarative-agent.json` and `foundry-plugin.json`
5. Publish to your organisation

**Via Teams Toolkit (VS Code):**
1. Create a new Teams app project
2. Replace the generated manifests with the files in this folder
3. Deploy using the Teams Toolkit sidebar

### 4. Test

1. Open Teams or M365 Chat
2. Start a conversation with the "HR Policy Assistant" agent
3. Ask a question about HR policies
4. The agent will search SharePoint and optionally call the Foundry plugin

## How OBO Works Here

In this pattern, **you don't manage OBO at all**. M365 Copilot handles the entire identity chain:

1. User is already authenticated in Teams/M365
2. When the declarative agent triggers a SharePoint search (via `GraphConnectors`), Copilot performs the OBO token exchange internally
3. SharePoint returns only content the user can access
4. If the plugin action is triggered, Copilot passes the retrieved context to your endpoint
5. Your endpoint uses a **service identity** (not the user's) to call Foundry — the content is already permission-trimmed

This is the most seamless OBO experience — but it only works within the M365 Copilot environment.

## Customising the Manifest

### Change SharePoint Scope

To restrict which SharePoint sites the agent searches, modify the `GraphConnectors` capability:
```json
"capabilities": [
  {
    "name": "GraphConnectors",
    "connections": [
      {
        "connection_id": "sharepoint",
        "sites": ["https://contoso.sharepoint.com/sites/HR"]
      }
    ]
  }
]
```

### Add More Actions

Add additional API plugins for different Foundry agents or external services:
```json
"actions": [
  { "id": "foundryAnalysis", "file": "foundry-plugin.json" },
  { "id": "ticketCreation", "file": "ticket-plugin.json" }
]
```

## Limitations

- **M365 Copilot only** — this pattern doesn't work as a standalone app
- **Limited orchestration control** — Copilot decides when to search SharePoint and when to call plugins
- **Plugin latency** — round-trip to your Foundry endpoint adds latency
- **Manifest schema changes** — the declarative agent schema is evolving; check the latest docs
- **Admin deployment** — publishing to an organisation requires Teams admin approval

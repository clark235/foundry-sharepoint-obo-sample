# Pattern 2: M365 Copilot Retrieval API + Foundry Agent

> **Approach:** You control the full pipeline. Acquire the user's token, call the Graph Retrieval API for permission-trimmed SharePoint content, then feed it to a Foundry agent for synthesis. Maximum control, maximum auditability.

---

## How It Works

```mermaid
%%{init: {'theme': 'dark', 'themeVariables': {'primaryColor': '#1e3a5f', 'primaryTextColor': '#e2e8f0', 'primaryBorderColor': '#3b82f6', 'lineColor': '#64748b'}}}%%
flowchart LR
    U["👤 User"] -->|"1. Sign in\n(MSAL)"| APP["🖥️ Your App"]

    subgraph Retrieval["Step 2: Retrieve — User's identity"]
        APP -->|"User's delegated token"| GRAPH["📊 Graph API\nRetrieval endpoint"]
        GRAPH -->|"Searches with\nuser's permissions"| SP[("📄 SharePoint")]
        SP -->|"Permission-trimmed\nchunks"| GRAPH
        GRAPH -->|"Trimmed content"| APP
    end

    subgraph Synthesis["Step 3: Synthesise — Service identity"]
        APP -->|"Pre-trimmed content\n+ user query"| FA["🤖 Foundry Agent"]
        FA -->|"Analysis result"| APP
    end

    APP -->|"4. Combined answer"| U

    style U fill:#1e3a5f,stroke:#3b82f6,color:#e2e8f0
    style APP fill:#713f12,stroke:#eab308,color:#fefce8
    style GRAPH fill:#713f12,stroke:#eab308,color:#fefce8
    style SP fill:#14532d,stroke:#22c55e,color:#f0fdf4
    style FA fill:#312e81,stroke:#6366f1,color:#eef2ff
```

## The Trust Model

Two identity boundaries keep the system secure:

```mermaid
%%{init: {'theme': 'dark', 'themeVariables': {'primaryColor': '#1e3a5f', 'primaryTextColor': '#e2e8f0', 'primaryBorderColor': '#3b82f6', 'lineColor': '#64748b'}}}%%
flowchart TB
    subgraph Boundary1["🔐 Boundary 1: User Identity"]
        direction LR
        U["👤 User token\n(delegated)"] --> GRAPH["Graph Retrieval API"]
        GRAPH --> SP[("SharePoint")]
        SP -->|"Only docs user\ncan access"| RESULT["Trimmed content ✅"]
    end

    subgraph Boundary2["🔐 Boundary 2: Service Identity"]
        direction LR
        SVC["🔑 Service credential\n(DefaultAzureCredential)"] --> FOUNDRY["Foundry Agent"]
        CONTENT["Pre-approved content\nfrom Boundary 1"] --> FOUNDRY
        FOUNDRY --> ANSWER["Synthesised answer"]
    end

    RESULT -->|"Content already safe\nto process"| CONTENT

    style Boundary1 fill:#1e3a5f,stroke:#3b82f6,color:#e2e8f0
    style Boundary2 fill:#312e81,stroke:#6366f1,color:#eef2ff
    style RESULT fill:#14532d,stroke:#22c55e,color:#f0fdf4
    style CONTENT fill:#14532d,stroke:#22c55e,color:#f0fdf4
```

**Why two boundaries?**
- **Boundary 1** ensures SharePoint enforces the user's permissions *before* content reaches your code
- **Boundary 2** lets Foundry process content without needing the user's token — it's already been vetted
- **Your app** is the bridge — it sees trimmed content, logs it if needed, and forwards to Foundry

---

## Prerequisites

| Requirement | Details |
|---|---|
| Azure AI Foundry project | With a pre-created agent (note the agent ID) |
| M365 Copilot licence | Or pay-as-you-go Copilot metering on the tenant |
| Azure AD app registration | Public client with delegated `Files.Read.All` + `Sites.Read.All` |
| Python 3.9+ | `pip install azure-ai-projects azure-identity msal requests` |

---

## Setup

### 1. Create Azure AD App Registration (Public Client)

```
Azure Portal → App registrations → + New registration
├── Name: "Foundry Retrieval Sample"
├── Account type: Single tenant
├── Redirect URI: http://localhost (Mobile/desktop)
├── Authentication → Allow public client flows: Yes
└── API permissions:
    ├── Microsoft Graph → Delegated → Files.Read.All
    └── Microsoft Graph → Delegated → Sites.Read.All
```

> **Admin consent:** Required for org-wide use. Individual user consent works for testing.

### 2. Create a Foundry Agent

Create an agent in your project (portal or SDK) and note its ID: `asst_xxxxxxxxxxxxxxxxxxxx`

### 3. Configure & Run

```bash
cp .env.example .env
# Fill in your values

pip install -r requirements.txt
python main.py
# → Browser opens for MSAL sign-in
# → Retrieval API returns SharePoint chunks
# → Foundry agent synthesises the answer
```

---

## Retrieval API Reference

```mermaid
%%{init: {'theme': 'dark', 'themeVariables': {'primaryColor': '#1e3a5f', 'primaryTextColor': '#e2e8f0', 'primaryBorderColor': '#3b82f6', 'lineColor': '#64748b'}}}%%
graph TB
    subgraph API["POST graph.microsoft.com/beta/copilot/retrieval"]
        direction TB
        REQ["Request"]
        REQ --> QS["queryString: 'remote work policy'"]
        REQ --> DS["dataSource: 'sharepoint'"]
        REQ --> MAX["maximumNumberOfResults: 5"]
        REQ --> FLT["filterExpression:\n'path:https://contoso.sharepoint.com/sites/HR'\n(optional — scope to specific sites)"]

        RESP["Response"]
        RESP --> HITS["retrievalHits[]"]
        HITS --> CONTENT["content.value: 'Chunk text...'"]
        HITS --> URL["resource.webUrl: 'https://...'"]
    end

    style API fill:#1e293b,stroke:#64748b,color:#94a3b8
    style REQ fill:#713f12,stroke:#eab308,color:#fefce8
    style RESP fill:#14532d,stroke:#22c55e,color:#f0fdf4
```

| Constraint | Value |
|---|---|
| **Endpoint** | `POST https://graph.microsoft.com/beta/copilot/retrieval` |
| **Auth** | Delegated user token (OBO) |
| **Rate limit** | 200 requests / user / hour |
| **Data sources** | `sharepoint`, `onedrive`, `teams` |
| **Returns** | Text chunks (not full documents) |
| **Status** | Beta — may change |

---

## When to Use This Pattern

```mermaid
%%{init: {'theme': 'dark', 'themeVariables': {'primaryColor': '#1e3a5f', 'primaryTextColor': '#e2e8f0', 'primaryBorderColor': '#3b82f6', 'lineColor': '#64748b'}}}%%
flowchart TB
    subgraph Good["✅ Good fit"]
        G1["Custom retrieval pipelines"]
        G2["Audit trail / compliance logging"]
        G3["Mixing SharePoint with other data sources"]
        G4["Full control over what the agent sees"]
        G5["Standalone apps — not in M365 Copilot"]
    end

    subgraph Bad["❌ Not a fit"]
        B1["Simple Q&A — overkill"]
        B2["High-volume queries\n(200/user/hour limit)"]
        B3["Need full documents\n(API returns chunks only)"]
    end

    style Good fill:#14532d,stroke:#22c55e,color:#f0fdf4
    style Bad fill:#7c2d12,stroke:#ef4444,color:#fef2f2
```

---

## Limitations

- **Beta API** — the Retrieval API is `beta` and subject to change
- **Rate limited** — 200 requests per user per hour; not suitable for high-frequency polling
- **Chunked results** — returns text snippets, not full documents
- **Interactive sign-in** — user must authenticate via browser (or you manage refresh tokens)
- **Two auth systems** — you manage MSAL for Graph *and* DefaultAzureCredential for Foundry

---

## Files

| File | Description |
|---|---|
| `main.py` | Complete pipeline: MSAL → Retrieval API → Foundry agent |
| `.env.example` | Environment variable template |
| `requirements.txt` | `azure-ai-projects`, `azure-identity`, `msal`, `requests` |

# Pattern 1: Foundry Agent with Native SharePoint Tool

> **Approach:** Use the built-in `SharepointToolDefinition` in Azure AI Foundry. The SDK handles OBO token exchange automatically — you write zero auth code for SharePoint.

---

## How It Works

```mermaid
%%{init: {'theme': 'dark', 'themeVariables': {'primaryColor': '#1e3a5f', 'primaryTextColor': '#e2e8f0', 'primaryBorderColor': '#3b82f6', 'lineColor': '#64748b'}}}%%
flowchart LR
    U["👤 User"] -->|"1. Signs in\n(browser popup)"| APP["🖥️ Your App\nAIProjectClient"]
    APP -->|"2. Creates agent with\nSharepointToolDefinition"| FA["🤖 Foundry Agent"]
    FA -->|"3. OBO token exchange\n(automatic)"| SP[("📄 SharePoint")]
    SP -->|"4. Permission-trimmed\ndocuments"| FA
    FA -->|"5. Answer + citations"| APP
    APP -->|"6. Display"| U

    style U fill:#1e3a5f,stroke:#3b82f6,color:#e2e8f0
    style APP fill:#312e81,stroke:#6366f1,color:#eef2ff
    style FA fill:#312e81,stroke:#6366f1,color:#eef2ff
    style SP fill:#14532d,stroke:#22c55e,color:#f0fdf4
```

## OBO Token Flow (Under the Hood)

```mermaid
%%{init: {'theme': 'dark', 'themeVariables': {'primaryColor': '#1e3a5f', 'primaryTextColor': '#e2e8f0', 'primaryBorderColor': '#3b82f6', 'lineColor': '#64748b'}}}%%
sequenceDiagram
    participant U as 👤 User
    participant App as 🖥️ Your App
    participant Entra as 🔐 Entra ID
    participant Foundry as 🤖 Foundry Agent Service
    participant SP as 📄 SharePoint

    U->>App: 1. Sign in via browser
    App->>Entra: 2. InteractiveBrowserCredential
    Entra-->>App: 3. User access token

    App->>Foundry: 4. Create agent + send message
    Note over App,Foundry: User token flows via AIProjectClient

    Foundry->>Entra: 5. OBO token exchange
    Note over Foundry,Entra: "Give me a SharePoint-scoped\ntoken for this user"
    Entra-->>Foundry: 6. SharePoint-scoped user token

    Foundry->>SP: 7. Search with user's identity
    SP-->>Foundry: 8. Only docs user can access

    Foundry-->>App: 9. Answer with citations
```

**The key insight:** You never touch SharePoint APIs, Graph tokens, or OBO logic. The `SharepointToolDefinition` handles all of it inside the Foundry service.

---

## Prerequisites

| Requirement | Details |
|---|---|
| Azure AI Foundry project | [Create at ai.azure.com](https://ai.azure.com) |
| M365 Copilot licence | On the user account — required for SharePoint grounding |
| SharePoint connection | Configured in Foundry project (see setup) |
| Azure AD app registration | Public client for interactive sign-in |
| Python 3.9+ | `pip install azure-ai-projects azure-identity` |

---

## Setup

### 1. Create Azure AD App Registration

```
Azure Portal → App registrations → + New registration
├── Name: "Foundry SharePoint Sample"
├── Account type: Single tenant
├── Redirect URI: http://localhost (Mobile/desktop)
└── Note: Client ID + Tenant ID
```

### 2. Create SharePoint Connection in Foundry

```
ai.azure.com → Your Project → Settings → Connected resources
├── + New connection → SharePoint
├── Authenticate with an M365 account
└── Note the Connection ID
```

> **Connection ID format:** `/subscriptions/<sub>/resourceGroups/<rg>/providers/Microsoft.CognitiveServices/accounts/<acct>/projects/<proj>/connections/<name>`

### 3. Configure & Run

```bash
cp .env.example .env
# Fill in your values

pip install -r requirements.txt
python main.py
# → Browser opens for sign-in
# → Agent queries SharePoint and prints results
```

---

## When to Use This Pattern

```mermaid
%%{init: {'theme': 'dark', 'themeVariables': {'primaryColor': '#1e3a5f', 'primaryTextColor': '#e2e8f0', 'primaryBorderColor': '#3b82f6', 'lineColor': '#64748b'}}}%%
flowchart TB
    subgraph Good["✅ Good fit"]
        G1["Internal tools and notebooks"]
        G2["Dev/test and prototyping"]
        G3["Custom AI model required"]
        G4["Multi-tool agents\n(AI Search + SharePoint + Code)"]
    end

    subgraph Bad["❌ Not a fit"]
        B1["Publishing to M365 Copilot\n(identity chain breaks)"]
        B2["Server-to-server / batch jobs\n(no interactive sign-in)"]
        B3["Users without M365 Copilot licence"]
    end

    style Good fill:#14532d,stroke:#22c55e,color:#f0fdf4
    style Bad fill:#7c2d12,stroke:#ef4444,color:#fef2f2
```

---

## Limitations

- **Cannot publish to M365 Copilot** — when surfaced through M365 Copilot (Teams), the user token doesn't flow through the M365 → Foundry boundary. Use [Pattern 3](../03-declarative-agent-manifest/) for M365 deployment.
- **Interactive sign-in required** — the user must authenticate via browser. No headless/service token scenario.
- **M365 Copilot licence required** — SharePoint tool returns empty results without it.
- **Project-level connection** — all agents in the same Foundry project share the SharePoint connection.

---

## Files

| File | Description |
|---|---|
| `main.py` | Complete working example — agent creation, query, response |
| `.env.example` | Environment variable template |
| `requirements.txt` | `azure-ai-projects`, `azure-identity` |

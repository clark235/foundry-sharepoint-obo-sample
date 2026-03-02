# Setup Guide — Declarative Agent with Foundry Plugin

## Architecture Overview

```mermaid
%%{ init: { 'theme': 'dark' } }%%
graph TB
    subgraph Dev["🛠️ Development"]
        MAN["declarative-agent.json\nAgent manifest"]
        PLUG["foundry-plugin.json\nOpenAPI plugin spec"]
        CODE["endpoint/app.py\nFlask → Foundry SDK"]
    end

    subgraph Deploy["🚀 Deployment"]
        TDP["Teams Developer Portal\ndev.teams.microsoft.com"]
        AAS["Azure App Service\nor Container App"]
    end

    subgraph Runtime["⚡ Runtime"]
        M365["M365 Copilot\nTeams / M365 Chat"]
        SP["SharePoint\nPermission-trimmed"]
        FA["Azure AI Foundry\nAgent + Model"]
    end

    MAN --> TDP
    PLUG --> TDP
    CODE --> AAS
    TDP --> M365
    AAS --> FA
    M365 -->|OBO| SP
    M365 -->|HTTP POST| AAS

    classDef dev fill:#2d4a6e,stroke:#4a7ab5,color:#fff
    classDef deploy fill:#1a5276,stroke:#2e86c1,color:#fff
    classDef runtime fill:#0d3b1e,stroke:#1e8449,color:#fff
    class MAN,PLUG,CODE dev
    class TDP,AAS deploy
    class M365,SP,FA runtime
```

---

## Step 1 — Create an Azure AI Foundry Agent

Before deploying anything, create the Foundry agent that will power your plugin endpoint.

```mermaid
%%{ init: { 'theme': 'dark' } }%%
sequenceDiagram
    participant You
    participant Portal as Foundry Portal\nai.azure.com
    participant Agent as Foundry Agent Service
    participant Model as gpt-4o Deployment

    You->>Portal: Navigate to your project
    Portal->>You: Project dashboard
    You->>Portal: Create new agent
    Portal->>Agent: Provision agent
    Agent->>Model: Link to model deployment
    Agent-->>You: Agent ID (asst_xxxx...)
    Note over You: Save this Agent ID for .env
```

**In the Foundry portal:**
1. Go to [ai.azure.com](https://ai.azure.com) → your project
2. Click **Agents** → **New agent**
3. Select model: `gpt-4o` or `gpt-4.1`
4. Set instructions:
   ```
   You are an expert analyst. You will receive a user question and SharePoint document
   context retrieved by M365 Copilot. Provide a clear, structured analysis with
   specific references to the provided content.
   ```
5. Save and copy the **Agent ID** (`asst_xxxxxxxxxxxxxxxxxxxx`)

---

## Step 2 — Deploy the Plugin Endpoint

The endpoint is a lightweight Flask app that bridges M365 Copilot to your Foundry agent.

```mermaid
%%{ init: { 'theme': 'dark' } }%%
flowchart LR
    subgraph Request["📥 Incoming from M365 Copilot"]
        REQ["POST /analyze\n{\n  query: string,\n  context: string\n}"]
    end

    subgraph Endpoint["🌐 Your Endpoint (endpoint/app.py)"]
        AUTH["DefaultAzureCredential\n(Managed Identity in prod)"]
        CLIENT["AIProjectClient\nFoundry SDK"]
        THREAD["Create thread\nCreate message\nRun agent"]
    end

    subgraph Response["📤 Response to M365 Copilot"]
        RESP["200 OK\n{\n  analysis: string,\n  confidence: number\n}"]
    end

    REQ --> AUTH
    AUTH --> CLIENT
    CLIENT --> THREAD
    THREAD --> RESP

    classDef req fill:#1a3a1a,stroke:#2e7d2e,color:#fff
    classDef ep fill:#1a2a4a,stroke:#2e5a9e,color:#fff
    classDef resp fill:#3a1a1a,stroke:#8e2020,color:#fff
    class REQ req
    class AUTH,CLIENT,THREAD ep
    class RESP resp
```

### Deploy to Azure App Service

```bash
cd endpoint

# Create App Service
az webapp up \
  --name your-foundry-plugin-endpoint \
  --runtime PYTHON:3.11 \
  --sku B2

# Set environment variables
az webapp config appsettings set \
  --name your-foundry-plugin-endpoint \
  --settings \
    FOUNDRY_PROJECT_ENDPOINT="https://<account>.services.ai.azure.com/api/projects/<project>" \
    FOUNDRY_AGENT_ID="asst_xxxxxxxxxxxxxxxxxxxx"

# Enable system-assigned managed identity (for production auth)
az webapp identity assign --name your-foundry-plugin-endpoint

# Grant the identity Azure AI User role on your Foundry account
PRINCIPAL_ID=$(az webapp identity show --name your-foundry-plugin-endpoint --query principalId -o tsv)
az role assignment create \
  --assignee $PRINCIPAL_ID \
  --role "Azure AI User" \
  --scope "/subscriptions/<sub>/resourceGroups/<rg>/providers/Microsoft.CognitiveServices/accounts/<account>"
```

### Update `foundry-plugin.json`

Replace the server URL with your deployed endpoint:
```json
"servers": [
  { "url": "https://your-foundry-plugin-endpoint.azurewebsites.net" }
]
```

---

## Step 3 — Register in Teams Developer Portal

```mermaid
%%{ init: { 'theme': 'dark' } }%%
flowchart TD
    START(["Open dev.teams.microsoft.com"])
    A["Apps → New App\nFill in name, description, icons"]
    B["Copilot agents section\n→ Add declarative agent"]
    C["Upload declarative-agent.json"]
    D["Add action\n→ Upload foundry-plugin.json"]
    E["Preview in M365 Copilot\nTest before publishing"]
    F{"Ready to\npublish?"}
    G["Publish → Submit for admin approval\nOrg-wide deployment"]
    H["Test in Teams\nM365 Chat / Outlook"]

    START --> A --> B --> C --> D --> E --> F
    F -->|Yes| G
    F -->|No, needs tweaks| C
    G --> H

    classDef action fill:#1a3a5e,stroke:#2e6bb5,color:#fff
    classDef decision fill:#3a2a0a,stroke:#b5862e,color:#fff
    classDef terminal fill:#0d3b1e,stroke:#1e8449,color:#fff
    class A,B,C,D,E,G,H action
    class F decision
    class START terminal
```

**Steps:**
1. Go to [Teams Developer Portal](https://dev.teams.microsoft.com/)
2. **Apps** → **New App**
3. Fill in basic info (name: "HR Policy Assistant", description, icons)
4. Navigate to **Copilot agents** in the left sidebar
5. Click **Add a declarative agent** → upload `declarative-agent.json`
6. Under **Actions**, click **Add action** → upload `foundry-plugin.json`
7. Click **Preview in Copilot** to test
8. When ready: **Publish** → **Submit for admin approval**

---

## Step 4 — How OBO Works in This Pattern

You don't write any OBO code — M365 Copilot handles it entirely.

```mermaid
%%{ init: { 'theme': 'dark' } }%%
sequenceDiagram
    participant User as 👤 User\n(signed into Teams)
    participant Copilot as 🤖 M365 Copilot\nDeclarative Agent
    participant Graph as Microsoft Graph\nCopilot Retrieval
    participant SP as 📂 SharePoint
    participant Endpoint as 🌐 Your Endpoint
    participant Foundry as ⚡ Foundry Agent

    User->>Copilot: "What does our remote work policy say?"

    Note over Copilot: Declarative agent instructions\ntrigger SharePoint search

    Copilot->>Graph: Search SharePoint\n(user's OBO token — automatic)
    Graph->>SP: Query with user's identity
    SP-->>Graph: Permission-trimmed results\n(only what user can access)
    Graph-->>Copilot: Document chunks + citations

    Note over Copilot: Needs deeper analysis?\nCall Foundry plugin action

    Copilot->>Endpoint: POST /analyze\n{query, context: <SP chunks>}

    Note over Endpoint: Content already permission-trimmed\nby SharePoint. Safe to process.

    Endpoint->>Foundry: Create thread + message\nRun agent (service identity)
    Foundry-->>Endpoint: Analysis result

    Endpoint-->>Copilot: {analysis, confidence}
    Copilot-->>User: Answer + SharePoint citations
```

**Key points:**
- **You never see the user's token** — M365 Copilot handles the entire OBO exchange
- **Permission trimming is automatic** — SharePoint enforces the user's access before content leaves
- **Your endpoint uses a service identity** — Managed Identity for Foundry calls (not the user's identity)

---

## Step 5 — Customise the Manifest

### Restrict to specific SharePoint sites

```json
"capabilities": [
  {
    "name": "GraphConnectors",
    "connections": [
      {
        "connection_id": "sharepoint",
        "sites": [
          "https://contoso.sharepoint.com/sites/HR",
          "https://contoso.sharepoint.com/sites/Legal"
        ]
      }
    ]
  }
]
```

### Add multiple Foundry actions

```json
"actions": [
  { "id": "foundryAnalysis",   "file": "foundry-plugin.json" },
  { "id": "ticketCreation",    "file": "ticket-plugin.json"  },
  { "id": "complianceCheck",   "file": "compliance-plugin.json" }
]
```

### Tune agent instructions

```json
"instructions": "You help employees with HR and legal policy questions. Always search SharePoint first and cite specific documents. For comparisons between policies or industry benchmarks, use the FoundryAnalysis action — do not attempt comparisons yourself. Never guess at policy content; if uncertain, say so and suggest the user consult HR directly."
```

---

## Troubleshooting

```mermaid
%%{ init: { 'theme': 'dark' } }%%
flowchart TD
    ERR["❌ Problem"]

    ERR --> Q1{"SharePoint\nnot returning docs?"}
    Q1 -->|Yes| A1["Check:\n1. User has M365 Copilot licence\n2. GraphConnectors capability in manifest\n3. User has access to the SP site"]

    ERR --> Q2{"Plugin action\nnot being called?"}
    Q2 -->|Yes| A2["Check:\n1. Endpoint URL in foundry-plugin.json\n2. App is running and reachable\n3. Agent instructions say when to call action"]

    ERR --> Q3{"Foundry agent\nnot responding?"}
    Q3 -->|Yes| A3["Check:\n1. FOUNDRY_PROJECT_ENDPOINT env var\n2. FOUNDRY_AGENT_ID is correct\n3. Managed identity has Azure AI User role"]

    ERR --> Q4{"403 on\nFoundry API?"}
    Q4 -->|Yes| A4["Run:\naz webapp identity assign\naz role assignment create\n--role 'Azure AI User'"]

    classDef error fill:#4a1a1a,stroke:#9e2020,color:#fff
    classDef fix fill:#1a3a1a,stroke:#2e7d2e,color:#fff
    class ERR error
    class A1,A2,A3,A4 fix
```

---

## Security Considerations

| Layer | Control | Notes |
|---|---|---|
| User → M365 Copilot | M365 SSO | Standard enterprise auth |
| SharePoint access | OBO (automatic) | Permission-trimmed by Graph |
| M365 Copilot → Endpoint | HTTPS + optional API key | Add `securitySchemes` to OpenAPI spec for production |
| Endpoint → Foundry | Managed Identity | No secrets stored in code |
| Foundry Agent | Azure RBAC | Azure AI User role scoped to project |

> **Never** store credentials in `app.py` or environment variables in development. Use Azure Key Vault for production secrets and Managed Identity for all Azure service calls.

---

*Azure AI Foundry Team · March 2026*  
*Related: [Declarative Agents Docs](https://learn.microsoft.com/microsoft-365-copilot/extensibility/overview-declarative-agent) · [Foundry Agent Service](https://learn.microsoft.com/azure/ai-foundry/agents/overview)*

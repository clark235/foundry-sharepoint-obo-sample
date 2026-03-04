# Pattern 3: Complete Implementation Walkthrough

> **Declarative Agent + Azure AI Foundry API Plugin** — the lowest-complexity, production-ready pattern for combining M365 Copilot's native SharePoint access with custom AI analysis.

---

## 1. Overview

Pattern 3 lets you build a Copilot agent that:

- **Searches SharePoint natively** — M365 Copilot handles OBO tokens, permission trimming, and document retrieval. You write zero auth code for this.
- **Calls your Foundry agent for deep analysis** — when the user's question requires comparison, synthesis, or compliance checking, Copilot sends the SharePoint content to your REST endpoint, which runs a Foundry agent.

**Why this pattern?**

| Alternative | Drawback | Pattern 3 advantage |
|---|---|---|
| Build your own RAG pipeline | You own auth, chunking, embedding, permission trimming | M365 Copilot does all of this |
| Use Foundry agent alone (Pattern 2) | Must implement SharePoint OBO token exchange yourself | SharePoint access is built-in |
| Use M365 Copilot alone (no plugin) | Limited to what the base model can do with retrieved docs | Foundry agent adds custom logic |

---

## 2. Architecture Deep Dive

```mermaid
%%{init: {'theme': 'dark', 'themeVariables': {'primaryColor': '#1e3a5f', 'primaryTextColor': '#e2e8f0', 'primaryBorderColor': '#3b82f6', 'lineColor': '#64748b'}}}%%
sequenceDiagram
    participant U as 👤 User (Teams)
    participant DA as 🤖 Declarative Agent
    participant M365 as 🔐 M365 Copilot Runtime
    participant SP as 📄 SharePoint
    participant TV as 🔑 Token Store
    participant EP as ☁️ /analyze Endpoint
    participant AUTH as 🛡️ Entra Token Validation
    participant FA as 🤖 Foundry Agent

    U->>DA: "Compare our remote work and PTO policies"
    
    rect rgb(20, 83, 45)
        Note over DA,SP: SharePoint — handled by M365 (OBO)
        DA->>M365: Search SharePoint + OneDrive
        M365->>SP: User's delegated token (OBO)
        SP-->>M365: Permission-trimmed docs ✅
        M365-->>DA: Policy documents
    end

    DA->>DA: Instructions say "comparison → use FoundryAnalysis"

    rect rgb(49, 46, 129)
        Note over DA,FA: Plugin — OAuth2 service-to-service
        DA->>TV: Request access token for plugin endpoint
        TV-->>DA: Bearer token
        DA->>EP: POST /analyze {query, context, user_display_name}
        EP->>AUTH: Validate bearer token (kid, aud, iss, exp)
        AUTH-->>EP: ✅ Valid
        EP->>FA: Create thread → send message → run agent
        FA-->>EP: Analysis result
        EP->>FA: Delete thread (cleanup)
        EP-->>DA: {analysis, sources_used, run_id}
    end

    DA-->>U: "Here's how the policies interact..."<br/>📎 Remote-Work-Policy.docx<br/>📎 PTO-Guidelines.docx
```

### Key points

1. **M365 Copilot handles ALL SharePoint auth** — OBO token exchange, permission trimming, and document retrieval happen automatically via the `OneDriveAndSharePoint` capability.
2. **Your endpoint receives a service-to-service token** — not the user's token. The token is issued via the OAuth registration in the Teams Developer Portal.
3. **The Foundry agent gets no user token** — it processes content using its own Managed Identity credentials. The user's display name is passed as plain text (optional).
4. **Thread lifecycle is per-request** — create, use, cleanup. No persistent threads.

---

## 3. Prerequisites Checklist

- [ ] **Microsoft 365 Copilot licence** — per user who will use the agent
- [ ] **Azure subscription** — for App Service / Container Apps + Foundry project
- [ ] **Azure AI Foundry project** — with an agent created and its ID noted
- [ ] **Entra ID app registration** — for the plugin endpoint (client ID + tenant ID)
- [ ] **Teams Developer Portal access** — [dev.teams.microsoft.com](https://dev.teams.microsoft.com)
- [ ] **Python 3.11+** — for local development
- [ ] **Azure CLI** — for deployment (`az webapp up` or `az containerapp up`)
- [ ] **Teams Toolkit or manual packaging** — to deploy the agent to Teams

---

## 4. Step 1 — Create the Foundry Agent

### Option A: Azure AI Foundry Portal

1. Go to [ai.azure.com](https://ai.azure.com) → your project
2. Select **Agents** → **Create agent**
3. Configure:
   - **Model**: GPT-4o (or your preferred model)
   - **Instructions**: Tailor for HR policy analysis — e.g.:
     ```
     You are an HR policy analyst. You receive a user's question and 
     SharePoint document content. Analyse the content to answer the 
     question. Focus on accuracy, cite specific sections, and identify 
     any gaps or conflicts between policies.
     ```
4. Note the **Agent ID** (starts with `asst_`)

### Option B: Python SDK

```python
from azure.ai.projects import AIProjectClient
from azure.identity import DefaultAzureCredential

client = AIProjectClient(
    endpoint="https://<account>.services.ai.azure.com/api/projects/<project>",
    credential=DefaultAzureCredential(),
)

agent = client.agents.create_agent(
    model="gpt-4o",
    name="HR Policy Analyst",
    instructions="""You are an HR policy analyst. You receive a user's question 
    and SharePoint document content. Analyse the content to answer the question. 
    Focus on accuracy, cite specific sections, and identify any gaps or conflicts 
    between policies.""",
)

print(f"Agent ID: {agent.id}")  # → asst_xxxxxxxxxxxxxxxxxxxx
```

---

## 5. Step 2 — Register the Entra ID App

This app registration protects your `/analyze` endpoint. M365 Copilot will authenticate against it.

1. Go to [Entra admin center](https://entra.microsoft.com) → **App registrations** → **New registration**
2. Name: `Foundry Analysis Endpoint`
3. Supported account types: **Single tenant** (your org only)
4. Register — note the **Application (client) ID** and **Tenant ID**
5. Under **Expose an API**:
   - Set Application ID URI: `api://<client-id>`
   - Add scope: `api://<client-id>/.default` (admin consent)
   - Add authorised client application: `ab3be6b7-f5df-413d-ac2d-abf1e3fd9c0b` (Microsoft's enterprise token store)
6. Under **Authentication**:
   - Add redirect URI: `https://teams.microsoft.com/api/platform/v1.0/oAuthConsentRedirect`

### Register in Teams Developer Portal

1. Go to [dev.teams.microsoft.com](https://dev.teams.microsoft.com) → **Tools** → **OAuth client registration**
2. Create registration:
   - **Base URL**: your endpoint URL (e.g., `https://foundry-endpoint.azurewebsites.net`)
   - **Client ID**: from step 4 above
   - **Authorization endpoint**: `https://login.microsoftonline.com/{tenant-id}/oauth2/v2.0/authorize`
   - **Token endpoint**: `https://login.microsoftonline.com/{tenant-id}/oauth2/v2.0/token`
   - **Scope**: `api://<client-id>/.default`
3. Note the **OAuth client registration ID** — this goes in the plugin manifest as `reference_id`

---

## 6. Step 3 — Deploy the Endpoint

### Option A: Azure App Service (quickest)

```bash
cd 03-declarative-agent-manifest/endpoint

# Deploy
az webapp up \
  --name foundry-endpoint \
  --runtime PYTHON:3.11 \
  --sku B1

# Configure environment
az webapp config appsettings set \
  --name foundry-endpoint \
  --settings \
    FOUNDRY_PROJECT_ENDPOINT="https://<acct>.services.ai.azure.com/api/projects/<proj>" \
    FOUNDRY_AGENT_ID="asst_xxxxxxxxxxxxxxxxxxxx" \
    AZURE_TENANT_ID="<tenant-id>" \
    AZURE_CLIENT_ID="<client-id>" \
    USE_MANAGED_IDENTITY="true"

# Grant managed identity access to Foundry
PRINCIPAL_ID=$(az webapp identity show --name foundry-endpoint --query principalId -o tsv)
az role assignment create \
  --assignee $PRINCIPAL_ID \
  --role "Azure AI Developer" \
  --scope /subscriptions/<sub>/resourceGroups/<rg>/providers/Microsoft.MachineLearningServices/workspaces/<foundry-project>
```

### Option B: Bicep (infrastructure as code)

```bash
az deployment group create \
  --resource-group <rg-name> \
  --template-file infra/main.bicep \
  --parameters \
    foundryProjectEndpoint='https://<acct>.services.ai.azure.com/api/projects/<proj>' \
    foundryAgentId='asst_xxxxxxxxxxxxxxxxxxxx' \
    tenantId='<tenant-id>' \
    clientId='<client-id>'
```

Then grant the managed identity role (see output `principalId`).

### Option C: Container Apps

```bash
cd 03-declarative-agent-manifest/endpoint

# Build and push image
az acr build --registry <your-acr> --image foundry-endpoint:latest .

# Deploy
az containerapp up \
  --name foundry-endpoint \
  --image <your-acr>.azurecr.io/foundry-endpoint:latest \
  --env-vars \
    FOUNDRY_PROJECT_ENDPOINT="..." \
    FOUNDRY_AGENT_ID="..." \
    AZURE_TENANT_ID="..." \
    AZURE_CLIENT_ID="..." \
    USE_MANAGED_IDENTITY="true"
```

### Verify deployment

```bash
curl https://foundry-endpoint.azurewebsites.net/health
# → {"foundry": "connected", "status": "ok"}
```

---

## 7. Step 4 — Configure the Manifests

After deployment, update the manifest files with your real values:

### `manifest/openapi.json`

Replace placeholder values:

```json
"servers": [
  { "url": "https://foundry-endpoint.azurewebsites.net" }
]
```

Update the OAuth security scheme with your tenant ID and client ID:

```json
"authorizationUrl": "https://login.microsoftonline.com/<tenant-id>/oauth2/v2.0/authorize",
"tokenUrl": "https://login.microsoftonline.com/<tenant-id>/oauth2/v2.0/token",
"scopes": {
  "api://<client-id>/.default": "Access the Foundry Analysis API"
}
```

### `manifest/foundry-plugin.json`

Set the `reference_id` to the OAuth registration ID from the Teams Developer Portal:

```json
"auth": {
  "type": "OAuthPluginVault",
  "reference_id": "<your-oauth-registration-id>"
}
```

### `manifest/declarative-agent.json`

Optionally scope SharePoint to specific sites:

```json
{
  "name": "OneDriveAndSharePoint",
  "items_by_url": [
    { "url": "https://contoso.sharepoint.com/sites/HR" }
  ]
}
```

### `manifest/app-manifest.json`

Update:
- `id` — your app's unique GUID
- `developer` — your organisation's details
- `validDomains` — your endpoint domain

---

## 8. Step 5 — Deploy to Teams

### Via Teams Developer Portal

1. Go to [dev.teams.microsoft.com](https://dev.teams.microsoft.com)
2. **Apps** → **New app**
3. Fill in basic info matching `app-manifest.json`
4. **Copilot agents** → **Add** → **Declarative agent**
5. Upload `declarative-agent.json`, `foundry-plugin.json`, and `openapi.json`
6. Add colour (192×192px) and outline (32×32px) icons
7. **Publish** → **Publish to your org**
8. A Teams admin approves in the **Teams admin centre**

### Via Agents Toolkit (VS Code)

1. Install the [Microsoft 365 Agents Toolkit](https://aka.ms/M365AgentsToolkit) extension
2. Create a new project → **Copilot Agent** → **Declarative Agent**
3. Replace the generated manifest files with yours
4. Press **F5** to sideload for testing, or use **Publish** for production

---

## 9. Step 6 — Test End to End

1. Open **Teams** or **M365 Chat**
2. Start a conversation with "HR Policy Assistant"
3. Try these prompts:

| Prompt | Expected behaviour |
|---|---|
| "What does our remote work policy say?" | SharePoint search only — direct answer with citations |
| "Compare our parental leave and PTO policies" | SharePoint search → Foundry plugin call → structured comparison |
| "Does working from Bali for 2 months comply with our remote work policy?" | SharePoint search → Foundry plugin call → compliance analysis |
| "What is the expense reimbursement limit?" | SharePoint search only — simple factual answer |

### Debugging tips

- Check **App Service logs**: `az webapp log tail --name foundry-endpoint`
- Check the `/health` endpoint first to confirm Foundry connectivity
- Each response includes a `run_id` — use this to find the run in the Foundry portal
- Use browser dev tools in Teams web to see the plugin API calls

---

## 10. Step 7 — Production Hardening

### Rate limiting

Add rate limiting to prevent abuse (e.g., using `flask-limiter`):

```python
from flask_limiter import Limiter
limiter = Limiter(app, default_limits=["60 per minute"])
```

### Monitoring

- Enable **Application Insights** on the App Service
- The endpoint logs every request with structured JSON — set up alerts on:
  - `foundry_timeout` errors (504s)
  - `foundry_error` errors (500s)
  - Latency > 30s

### Secret management

- Move `AZURE_CLIENT_ID` and other sensitive settings to **Azure Key Vault**
- Reference Key Vault secrets from App Service settings:
  ```
  @Microsoft.KeyVault(VaultName=my-vault;SecretName=client-id)
  ```

### Network isolation

- Restrict the App Service to accept traffic only from M365 Copilot IPs
- Enable **VNet integration** if the Foundry project is in a private network

---

## 11. Troubleshooting

### Problem 1: Agent never calls the Foundry plugin

**Symptom**: Copilot always answers from SharePoint, never invokes `/analyze`.

**Fix**: Improve the `instructions` in `declarative-agent.json`. Be explicit about when to call the plugin. The instructions must mention specific trigger phrases (e.g., "compare", "analyse", "comply", "synthesise").

### Problem 2: 401 Unauthorized from the endpoint

**Symptom**: Plugin call fails with `{"error": "unauthorized"}`.

**Fix**:
1. Verify the OAuth registration ID in `foundry-plugin.json` matches the Teams Developer Portal
2. Confirm the `AZURE_CLIENT_ID` in App Service settings matches the Entra app registration
3. Check that the enterprise token store client (`ab3be6b7-...`) is added as an authorised client application
4. Ensure the redirect URI `https://teams.microsoft.com/api/platform/v1.0/oAuthConsentRedirect` is registered

### Problem 3: Foundry agent times out (504)

**Symptom**: `/analyze` returns `{"error": "foundry_timeout"}`.

**Fix**:
1. Increase `FOUNDRY_TIMEOUT_SECONDS` (default: 90)
2. Increase gunicorn `--timeout` to match (in Dockerfile or App Service startup command)
3. Check if the Foundry agent model deployment has sufficient capacity
4. Simplify the agent's instructions to reduce processing time

### Problem 4: SharePoint returns no results

**Symptom**: Agent says "I couldn't find any relevant policies".

**Fix**:
1. Confirm the user has access to the SharePoint sites in question
2. If you scoped SharePoint in the manifest (`items_by_url`), verify the URLs are correct
3. Check that documents are indexed — newly uploaded files may take time
4. Try broader terms in your query

### Problem 5: "This app is blocked" in Teams

**Symptom**: Users can't access the agent.

**Fix**: The Teams admin needs to approve the app in the **Teams admin centre** → **Manage apps**. Custom apps may be blocked by org policy.

---

## 12. What M365 Copilot Actually Sends

When M365 Copilot decides to call your plugin, here's what the HTTP request looks like:

```http
POST https://foundry-endpoint.azurewebsites.net/analyze
Authorization: Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsImtpZCI6Ii1LSTNR...
Content-Type: application/json
X-Request-ID: 7b3f8a1e-2d4c-4e6f-9a8b-1c3d5e7f9a2b

{
  "query": "Compare our remote work policy with the PTO policy — how do they interact for employees working abroad?",
  "context": "## Remote Work Policy (HR-Policy-2025.docx)\n\nSection 4.2 — International Remote Work\nEmployees may work remotely from outside their home country for up to 30 calendar days per year...\n\n## PTO Guidelines (PTO-Guidelines-2025.docx)\n\nSection 2.1 — Accrual\nFull-time employees accrue 20 days of PTO per calendar year...\n\nSection 3.4 — PTO During International Assignments\nPTO taken during international remote work periods counts against both the PTO balance and the 30-day international remote work allowance...",
  "user_display_name": "Jane Doe"
}
```

And here's the response your endpoint returns:

```json
{
  "analysis": "The remote work and PTO policies interact in the following ways:\n\n1. **Shared cap**: PTO taken while working internationally counts against BOTH the PTO balance AND the 30-day international remote work allowance (Section 3.4 of PTO Guidelines).\n\n2. **Practical impact**: An employee using 10 PTO days abroad effectively uses 10 of their 30 international work days, leaving only 20 days for actual remote work abroad.\n\n3. **Recommendation**: Employees planning extended international work should use PTO for personal days in their home country to maximise their international remote work allowance.",
  "sources_used": 4,
  "run_id": "run_abc123def456"
}
```

---

## File Structure

```
03-declarative-agent-manifest/
├── manifest/
│   ├── declarative-agent.json     ← Agent manifest (name, instructions, capabilities, actions)
│   ├── foundry-plugin.json        ← API plugin manifest (M365 plugin format v2.2)
│   ├── openapi.json               ← OpenAPI 3.0 spec for /analyze and /health
│   └── app-manifest.json          ← Teams app package manifest
├── endpoint/
│   ├── app.py                     ← Flask application (routes, validation, logging)
│   ├── auth.py                    ← Entra ID bearer token validation
│   ├── foundry_client.py          ← Foundry agent thread lifecycle
│   ├── .env.example               ← Environment variable template
│   ├── requirements.txt           ← Python dependencies
│   └── Dockerfile                 ← Container build for Azure Container Apps
├── infra/
│   └── main.bicep                 ← Azure infrastructure (App Service + identity)
├── tests/
│   └── test_endpoint.py           ← Pytest tests with mocked Foundry calls
├── WALKTHROUGH.md                 ← This file
└── README.md                      ← Pattern overview and quick-start
```

"""
Pattern 2: Foundry Agent + M365 Copilot Retrieval API
User authenticates interactively → delegated token → Graph Retrieval API → SharePoint content
Then feeds that content to a Foundry agent for synthesis.

Requirements:
- pip install azure-ai-projects azure-identity msal requests
- M365 Copilot licence OR pay-as-you-go enabled on the tenant
- App registration with Files.Read.All, Sites.Read.All delegated permissions (user-consented)
"""

import os, requests
from azure.ai.projects import AIProjectClient
from azure.identity import InteractiveBrowserCredential
from msal import PublicClientApplication

TENANT_ID   = os.environ["AZURE_TENANT_ID"]
CLIENT_ID   = os.environ["AZURE_CLIENT_ID"]   # App registration (public client)
SCOPES      = ["https://graph.microsoft.com/Files.Read.All",
               "https://graph.microsoft.com/Sites.Read.All"]

# ── Step 1: Acquire delegated token for Graph API ─────────────────────────────
# This is the OBO flow — user signs in, their identity is used for SharePoint access
app = PublicClientApplication(CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}")
result = app.acquire_token_interactive(scopes=SCOPES)

if "access_token" not in result:
    raise Exception(f"Auth failed: {result.get('error_description')}")

user_token = result["access_token"]
print("User authenticated successfully")

# ── Step 2: Call M365 Copilot Retrieval API ────────────────────────────────────
# This API returns permission-trimmed SharePoint chunks for the signed-in user
# Only content the user can access is returned — OBO enforced by Microsoft Graph
query = "remote work policy working from another country"

retrieval_response = requests.post(
    "https://graph.microsoft.com/beta/copilot/retrieval",
    headers={
        "Authorization": f"Bearer {user_token}",
        "Content-Type": "application/json",
    },
    json={
        "queryString": query,
        "dataSource": "sharepoint",                     # or "onedrive", "teams"
        "maximumNumberOfResults": 5,
        # Optional: scope to specific site
        # "filterExpression": "path:https://contoso.sharepoint.com/sites/HR"
    },
)

if retrieval_response.status_code != 200:
    raise Exception(f"Retrieval API error {retrieval_response.status_code}: {retrieval_response.text}")

hits = retrieval_response.json().get("retrievalHits", [])
print(f"Retrieved {len(hits)} document chunks from SharePoint")

# Format retrieved content for the agent
context = "\n\n---\n\n".join([
    f"**Source:** {hit.get('resource', {}).get('webUrl', 'Unknown')}\n{hit.get('content', {}).get('value', '')}"
    for hit in hits
])

# ── Step 3: Send to Foundry agent for synthesis ───────────────────────────────
# The agent uses its own identity (not the user's) for Foundry operations
# SharePoint content is already permission-trimmed — safe to pass as context
from azure.identity import DefaultAzureCredential

foundry_client = AIProjectClient(
    endpoint=os.environ["FOUNDRY_PROJECT_ENDPOINT"],
    credential=DefaultAzureCredential(),    # Agent identity for Foundry
)

thread = foundry_client.agents.create_thread()

foundry_client.agents.create_message(
    thread_id=thread.id,
    role="user",
    content=f"""The user asked: "{query}"

Here is relevant content retrieved from their SharePoint (they have permission to see this):

{context}

Please synthesise a clear, helpful answer based on this content.""",
)

run = foundry_client.agents.create_and_process_run(
    thread_id=thread.id,
    agent_id=os.environ["FOUNDRY_AGENT_ID"],  # Pre-created agent ID
)

messages = foundry_client.agents.list_messages(thread_id=thread.id)
for msg in messages:
    if msg.role == "assistant":
        print(f"\nAgent response:\n{msg.content[0].text.value}")

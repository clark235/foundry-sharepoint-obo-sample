"""
Pattern 1: Azure AI Foundry Agent with native SharePoint tool
Uses identity passthrough (OBO) — user's permissions enforced by SharePoint

Requirements:
- pip install azure-ai-projects azure-identity
- M365 Copilot licence on the user account
- Azure AI Foundry project with SharePoint connection configured
- Environment variables: see .env.example
"""

import os
from azure.ai.projects import AIProjectClient
from azure.ai.projects.models import (
    Agent,
    SharepointToolDefinition,
    ThreadMessage,
    MessageRole,
)
from azure.identity import DefaultAzureCredential, InteractiveBrowserCredential

# ── Auth ──────────────────────────────────────────────────────────────────────
# For local dev: InteractiveBrowserCredential prompts the user to sign in
# This is critical for OBO — the user's identity must flow through
# In production: use OnBehalfOfCredential or the Foundry SDK's built-in OBO
credential = InteractiveBrowserCredential(
    tenant_id=os.environ["AZURE_TENANT_ID"],
    client_id=os.environ["AZURE_CLIENT_ID"],
)

# ── Foundry client ────────────────────────────────────────────────────────────
client = AIProjectClient(
    endpoint=os.environ["FOUNDRY_PROJECT_ENDPOINT"],  # https://<account>.services.ai.azure.com/api/projects/<project>
    credential=credential,
)

# ── Create agent with SharePoint tool ────────────────────────────────────────
# SharepointToolDefinition uses identity passthrough automatically
# The user's OBO token is passed to SharePoint — only docs they can access are returned
sharepoint_tool = SharepointToolDefinition(
    sharepoint_connection_id=os.environ["SHAREPOINT_CONNECTION_ID"],
    # Connection ID format: /subscriptions/.../connections/<name>
    # Create this connection in Foundry portal: Project Settings → Connected resources → Add → SharePoint
)

agent: Agent = client.agents.create_agent(
    model=os.environ.get("FOUNDRY_MODEL_DEPLOYMENT", "gpt-4o"),
    name="SharePoint Assistant",
    instructions="""You are a helpful assistant with access to SharePoint documents.
    When answering questions, search SharePoint for relevant documents and cite your sources.
    Only reference documents the user has permission to access.""",
    tools=[sharepoint_tool],
)
print(f"Created agent: {agent.id}")

# ── Run a query ───────────────────────────────────────────────────────────────
thread = client.agents.create_thread()

client.agents.create_message(
    thread_id=thread.id,
    role=MessageRole.USER,
    content="What does our remote work policy say about working from a different country?",
)

run = client.agents.create_and_process_run(
    thread_id=thread.id,
    agent_id=agent.id,
)
print(f"Run status: {run.status}")

# ── Print response ─────────────────────────────────────────────────────────────
messages = client.agents.list_messages(thread_id=thread.id)
for msg in messages:
    if msg.role == MessageRole.ASSISTANT:
        print(f"\nAgent: {msg.content[0].text.value}")
        # Print citations if SharePoint returned grounded results
        if hasattr(msg.content[0].text, "annotations"):
            for ann in msg.content[0].text.annotations:
                print(f"  Source: {ann.text} → {getattr(ann, 'url', 'N/A')}")

# ── Cleanup ────────────────────────────────────────────────────────────────────
client.agents.delete_agent(agent.id)

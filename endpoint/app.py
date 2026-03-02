"""
Azure AI Foundry Plugin Endpoint
---------------------------------
Receives POST /analyze calls from M365 Copilot declarative agent,
runs the query + SharePoint context through a Foundry agent, and returns the analysis.

Deploy to: Azure App Service, Container Apps, or Azure Functions
Auth:       DefaultAzureCredential (Managed Identity in production, az login locally)

Environment variables (see .env.example):
    FOUNDRY_PROJECT_ENDPOINT   https://<account>.services.ai.azure.com/api/projects/<project>
    FOUNDRY_AGENT_ID           asst_xxxxxxxxxxxxxxxxxxxx
"""

import os
import logging
from flask import Flask, request, jsonify, abort
from azure.ai.projects import AIProjectClient
from azure.ai.projects.models import MessageRole
from azure.identity import DefaultAzureCredential
from azure.core.exceptions import HttpResponseError

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

app = Flask(__name__)

# ── Foundry client (initialised once at startup) ──────────────────────────────
# DefaultAzureCredential: uses Managed Identity in Azure, az login locally
_foundry_client: AIProjectClient | None = None

def get_client() -> AIProjectClient:
    global _foundry_client
    if _foundry_client is None:
        endpoint = os.environ.get("FOUNDRY_PROJECT_ENDPOINT")
        if not endpoint:
            raise EnvironmentError("FOUNDRY_PROJECT_ENDPOINT is not set")
        _foundry_client = AIProjectClient(
            endpoint=endpoint,
            credential=DefaultAzureCredential(),
        )
    return _foundry_client


# ── Health check ──────────────────────────────────────────────────────────────
@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


# ── Main plugin endpoint ──────────────────────────────────────────────────────
@app.route("/analyze", methods=["POST"])
def analyze():
    """
    Called by M365 Copilot when the declarative agent decides to invoke the
    FoundryAnalysis action. Receives the user's original query and permission-
    trimmed SharePoint context, runs it through the Foundry agent, and returns
    a structured analysis.

    Request body (from M365 Copilot):
        {
          "query":   "What does the remote work policy say about international travel?",
          "context": "<SharePoint document chunks — already permission-trimmed>"
        }

    Response:
        {
          "analysis":   "Based on the retrieved policies...",
          "confidence": 0.9
        }
    """
    data = request.get_json(silent=True)
    if not data:
        abort(400, "Request body must be JSON")

    query   = data.get("query", "").strip()
    context = data.get("context", "").strip()

    if not query:
        abort(400, "Missing required field: query")

    agent_id = os.environ.get("FOUNDRY_AGENT_ID")
    if not agent_id:
        abort(500, "FOUNDRY_AGENT_ID is not configured")

    log.info(f"Analyze request — query length: {len(query)}, context length: {len(context)}")

    try:
        client = get_client()

        # Create a fresh thread for each request
        # (stateless — each M365 Copilot turn is independent)
        thread = client.agents.create_thread()

        # Build the message:
        # - Context is already permission-trimmed by SharePoint/M365 Copilot
        # - Safe to pass directly to the Foundry agent
        message_content = f"""User question: {query}

SharePoint content retrieved by M365 Copilot (permission-trimmed for this user):

{context if context else "(No SharePoint context was retrieved — answer based on general knowledge if appropriate, or indicate that no relevant documents were found.)"}

Please provide a clear, structured analysis based on the above content. Cite specific
document sections where relevant. If the content does not contain enough information
to answer the question confidently, say so clearly."""

        client.agents.create_message(
            thread_id=thread.id,
            role=MessageRole.USER,
            content=message_content,
        )

        # Run the agent synchronously (create_and_process_run polls until done)
        run = client.agents.create_and_process_run(
            thread_id=thread.id,
            agent_id=agent_id,
        )

        if run.status == "failed":
            log.error(f"Agent run failed: {run.last_error}")
            return jsonify({
                "analysis":   "The analysis could not be completed due to an internal error.",
                "confidence": 0.0,
            }), 500

        # Extract the assistant's response
        messages = client.agents.list_messages(thread_id=thread.id)
        for msg in messages:
            if msg.role == MessageRole.ASSISTANT and msg.content:
                analysis_text = msg.content[0].text.value
                log.info(f"Analysis generated — {len(analysis_text)} chars")
                return jsonify({
                    "analysis":   analysis_text,
                    "confidence": 0.9,
                })

        return jsonify({
            "analysis":   "No analysis was generated. Please try rephrasing your question.",
            "confidence": 0.0,
        })

    except HttpResponseError as e:
        log.error(f"Foundry API error: {e.status_code} — {e.message}")
        abort(502, f"Foundry API error: {e.message}")
    except Exception as e:
        log.error(f"Unexpected error: {e}", exc_info=True)
        abort(500, "An unexpected error occurred")


if __name__ == "__main__":
    # Local development only — use gunicorn or Azure App Service in production
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port, debug=False)

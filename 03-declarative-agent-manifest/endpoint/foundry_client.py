"""
Manages Azure AI Foundry agent thread lifecycle for each request.

Flow per request:
  1. Create a thread
  2. Add a user message with query + SharePoint context
  3. Create and process a run (synchronous — blocks until the agent completes)
  4. Extract the assistant's response
  5. Delete the thread (cleanup)

Uses Managed Identity in production (Azure App Service / Container Apps)
and DefaultAzureCredential for local development.
"""

import logging
import os
import time

from azure.ai.projects import AIProjectClient
from azure.identity import DefaultAzureCredential, ManagedIdentityCredential

logger = logging.getLogger("foundry-endpoint.foundry_client")

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
FOUNDRY_PROJECT_ENDPOINT = os.environ.get("FOUNDRY_PROJECT_ENDPOINT", "")
FOUNDRY_AGENT_ID = os.environ.get("FOUNDRY_AGENT_ID", "")
FOUNDRY_TIMEOUT_SECONDS = int(os.environ.get("FOUNDRY_TIMEOUT_SECONDS", "90"))
USE_MANAGED_IDENTITY = os.environ.get("USE_MANAGED_IDENTITY", "false").lower() == "true"

# ---------------------------------------------------------------------------
# Client singleton
# ---------------------------------------------------------------------------
_client: AIProjectClient | None = None


def get_client() -> AIProjectClient:
    """
    Returns a cached AIProjectClient.

    Uses ManagedIdentityCredential on Azure, DefaultAzureCredential locally.
    """
    global _client
    if _client is not None:
        return _client

    if USE_MANAGED_IDENTITY:
        credential = ManagedIdentityCredential()
        logger.info("Using ManagedIdentityCredential for Foundry client")
    else:
        credential = DefaultAzureCredential()
        logger.info("Using DefaultAzureCredential for Foundry client")

    _client = AIProjectClient(
        endpoint=FOUNDRY_PROJECT_ENDPOINT,
        credential=credential,
    )
    return _client


# ---------------------------------------------------------------------------
# Agent execution
# ---------------------------------------------------------------------------
def run_agent(
    query: str,
    context: str,
    user_display_name: str | None = None,
) -> dict:
    """
    Run the Foundry agent with the given query and SharePoint context.

    Creates a thread, sends the message, waits for completion, extracts
    the response, and cleans up the thread.

    Args:
        query:             The user's original question.
        context:           SharePoint content retrieved by M365 Copilot.
        user_display_name: Optional display name for personalisation.

    Returns:
        dict with keys: analysis (str), sources_used (int), run_id (str)

    Raises:
        TimeoutError: if the Foundry agent doesn't complete in time.
        RuntimeError: if the agent run fails or returns no response.
    """
    client = get_client()
    thread = None

    try:
        # --- 1. Create thread -----------------------------------------------
        thread = client.agents.create_thread()
        logger.info(f"Created Foundry thread {thread.id}")

        # --- 2. Compose and send message ------------------------------------
        user_line = (
            f"User ({user_display_name}) asks:" if user_display_name
            else "User asks:"
        )
        message_content = (
            f"{user_line}\n{query}\n\n"
            f"--- SharePoint Context ---\n{context}"
        )

        client.agents.create_message(
            thread_id=thread.id,
            role="user",
            content=message_content,
        )

        # --- 3. Run the agent -----------------------------------------------
        start = time.time()
        run = client.agents.create_and_process_run(
            thread_id=thread.id,
            agent_id=FOUNDRY_AGENT_ID,
        )

        elapsed = time.time() - start
        if elapsed > FOUNDRY_TIMEOUT_SECONDS:
            raise TimeoutError(
                f"Foundry agent run took {elapsed:.1f}s "
                f"(limit: {FOUNDRY_TIMEOUT_SECONDS}s)"
            )

        if run.status != "completed":
            raise RuntimeError(
                f"Foundry agent run did not complete — status: {run.status}, "
                f"last_error: {getattr(run, 'last_error', 'unknown')}"
            )

        # --- 4. Extract assistant response -----------------------------------
        messages = client.agents.list_messages(thread_id=thread.id)
        analysis_text = None

        for msg in messages:
            if msg.role == "assistant" and msg.content:
                # The first assistant message is the agent's response
                analysis_text = msg.content[0].text.value
                break

        if not analysis_text:
            raise RuntimeError("Foundry agent returned no assistant message")

        # Count approximate source references in the context
        # (heuristic: count document separators or paragraphs)
        sources_used = max(1, context.count("\n\n"))

        logger.info(
            f"Foundry run {run.id} completed in {elapsed:.1f}s — "
            f"response length={len(analysis_text)}"
        )

        return {
            "analysis": analysis_text,
            "sources_used": sources_used,
            "run_id": run.id,
        }

    finally:
        # --- 5. Cleanup thread -----------------------------------------------
        if thread is not None:
            try:
                client.agents.delete_thread(thread_id=thread.id)
                logger.info(f"Deleted Foundry thread {thread.id}")
            except Exception as cleanup_exc:
                logger.warning(
                    f"Failed to delete Foundry thread {thread.id}: {cleanup_exc}"
                )

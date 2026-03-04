"""
Production endpoint for M365 Copilot → Azure AI Foundry Agent bridge.

Security: validates Bearer token from M365 Copilot via Entra ID token validation.
Foundry: creates a thread, sends context, processes response, cleans up.

Endpoints:
    GET  /health   — returns service health and Foundry connectivity
    POST /analyze  — runs Foundry agent analysis on SharePoint content
"""

import logging
import os
import sys
import time

from flask import Flask, request, jsonify

from auth import require_bearer_token
from foundry_client import get_client, run_agent

# ---------------------------------------------------------------------------
# Structured logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format='{"time":"%(asctime)s","level":"%(levelname)s","logger":"%(name)s","message":"%(message)s"}',
    stream=sys.stdout,
)
logger = logging.getLogger("foundry-endpoint")

# ---------------------------------------------------------------------------
# Required environment variables — fail fast if missing
# ---------------------------------------------------------------------------
REQUIRED_ENV = [
    "FOUNDRY_PROJECT_ENDPOINT",
    "FOUNDRY_AGENT_ID",
    "AZURE_TENANT_ID",
    "AZURE_CLIENT_ID",
]


def _validate_env():
    """Validate required environment variables at startup."""
    missing = [v for v in REQUIRED_ENV if not os.environ.get(v)]
    if missing:
        logger.error(f"Missing required environment variables: {', '.join(missing)}")
        sys.exit(1)
    logger.info("Environment validation passed")


_validate_env()

# ---------------------------------------------------------------------------
# Flask app
# ---------------------------------------------------------------------------
app = Flask(__name__)


@app.route("/health", methods=["GET"])
def health():
    """
    Health check endpoint for Azure App Service / Container Apps probes.
    Verifies that the Foundry client can be instantiated.
    """
    try:
        client = get_client()
        # Light check — just confirm the client initialises without error
        foundry_status = "connected" if client is not None else "unavailable"
    except Exception as exc:
        logger.warning(f"Health check: Foundry client init failed: {exc}")
        foundry_status = "unavailable"

    status_code = 200 if foundry_status == "connected" else 503
    return jsonify({"status": "ok" if foundry_status == "connected" else "degraded",
                     "foundry": foundry_status}), status_code


@app.route("/analyze", methods=["POST"])
@require_bearer_token
def analyze():
    """
    Analyse SharePoint content using Azure AI Foundry agent.

    Expected JSON body:
        query   (str, required)  — the user's original question
        context (str, required)  — SharePoint content retrieved by M365 Copilot
        user_display_name (str, optional) — user's display name for personalisation

    Returns:
        200: { analysis, sources_used, run_id }
        400: missing / invalid fields
        401: invalid bearer token (handled by decorator)
        500: Foundry agent error
    """
    start_time = time.time()
    request_id = request.headers.get("X-Request-ID", "unknown")

    # --- Request validation ------------------------------------------------
    if not request.is_json:
        logger.warning(f"[{request_id}] Non-JSON content type")
        return jsonify({"error": "invalid_content_type",
                        "message": "Content-Type must be application/json"}), 400

    data = request.get_json(silent=True)
    if data is None:
        logger.warning(f"[{request_id}] Malformed JSON body")
        return jsonify({"error": "invalid_json",
                        "message": "Request body must be valid JSON"}), 400

    query = data.get("query")
    context = data.get("context")
    user_display_name = data.get("user_display_name")

    if not query or not isinstance(query, str) or not query.strip():
        return jsonify({"error": "missing_field",
                        "message": "'query' is required and must be a non-empty string"}), 400

    if not context or not isinstance(context, str) or not context.strip():
        return jsonify({"error": "missing_field",
                        "message": "'context' is required and must be a non-empty string"}), 400

    if user_display_name is not None and not isinstance(user_display_name, str):
        return jsonify({"error": "invalid_field",
                        "message": "'user_display_name' must be a string if provided"}), 400

    logger.info(f"[{request_id}] /analyze called — query length={len(query)}, "
                f"context length={len(context)}, user={user_display_name or 'anonymous'}")

    # --- Call Foundry agent ------------------------------------------------
    try:
        result = run_agent(
            query=query.strip(),
            context=context.strip(),
            user_display_name=user_display_name,
        )
    except TimeoutError:
        elapsed = time.time() - start_time
        logger.error(f"[{request_id}] Foundry agent timed out after {elapsed:.1f}s")
        return jsonify({"error": "foundry_timeout",
                        "message": "The Foundry agent did not respond in time. Please try again."}), 504
    except Exception as exc:
        elapsed = time.time() - start_time
        logger.error(f"[{request_id}] Foundry agent error after {elapsed:.1f}s: {exc}")
        return jsonify({"error": "foundry_error",
                        "message": "An error occurred while processing your request. Please try again."}), 500

    elapsed = time.time() - start_time
    logger.info(f"[{request_id}] /analyze completed in {elapsed:.1f}s — "
                f"run_id={result.get('run_id', 'n/a')}")

    return jsonify(result), 200


# ---------------------------------------------------------------------------
# Dev server (production uses gunicorn)
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=True)

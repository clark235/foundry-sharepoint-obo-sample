"""
Validates Bearer tokens from M365 Copilot.

Tokens are issued by Microsoft Entra ID and scoped to the app registration
that protects this endpoint.  We validate:
  - Signature (via Entra OIDC JWKS)
  - Audience (must match AZURE_CLIENT_ID)
  - Issuer  (must match the tenant)
  - Expiry

Requires:
    AZURE_TENANT_ID   — your Entra tenant ID
    AZURE_CLIENT_ID   — the app registration client ID for this endpoint
"""

import logging
import os
import time
from functools import wraps

import jwt
import requests
from flask import request, jsonify

logger = logging.getLogger("foundry-endpoint.auth")

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
TENANT_ID = os.environ.get("AZURE_TENANT_ID", "")
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID", "")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
OIDC_CONFIG_URL = f"{AUTHORITY}/v2.0/.well-known/openid-configuration"

# Cache JWKS keys for 1 hour
_jwks_cache: dict = {"keys": None, "fetched_at": 0}
_JWKS_TTL = 3600  # seconds


# ---------------------------------------------------------------------------
# JWKS helpers
# ---------------------------------------------------------------------------
def _fetch_jwks() -> dict:
    """Fetch JSON Web Key Set from Entra OIDC metadata."""
    now = time.time()
    if _jwks_cache["keys"] and (now - _jwks_cache["fetched_at"]) < _JWKS_TTL:
        return _jwks_cache["keys"]

    logger.info("Fetching OIDC configuration and JWKS from Entra")
    oidc_resp = requests.get(OIDC_CONFIG_URL, timeout=10)
    oidc_resp.raise_for_status()
    jwks_uri = oidc_resp.json()["jwks_uri"]

    jwks_resp = requests.get(jwks_uri, timeout=10)
    jwks_resp.raise_for_status()
    keys = jwks_resp.json()

    _jwks_cache["keys"] = keys
    _jwks_cache["fetched_at"] = now
    return keys


def _get_signing_key(token: str):
    """Extract the signing key that matches the token's kid header."""
    jwks = _fetch_jwks()
    jwk_client = jwt.PyJWKClient.__new__(jwt.PyJWKClient)
    # Build a JWKSet from the fetched keys
    from jwt.api_jwk import PyJWKSet
    key_set = PyJWKSet.from_dict(jwks)

    unverified_header = jwt.get_unverified_header(token)
    kid = unverified_header.get("kid")
    if not kid:
        raise jwt.InvalidTokenError("Token header missing 'kid'")

    for key in key_set.keys:
        if key.key_id == kid:
            return key.key
    raise jwt.InvalidTokenError(f"No matching key found for kid={kid}")


# ---------------------------------------------------------------------------
# Token validation
# ---------------------------------------------------------------------------
def _validate_token(token: str) -> dict:
    """
    Validate a Bearer token from Entra ID.

    Returns the decoded claims on success, raises on failure.
    """
    signing_key = _get_signing_key(token)

    # Entra v2 tokens use issuer with the tenant ID
    expected_issuer = f"https://login.microsoftonline.com/{TENANT_ID}/v2.0"

    claims = jwt.decode(
        token,
        signing_key,
        algorithms=["RS256"],
        audience=CLIENT_ID,
        issuer=expected_issuer,
        options={
            "verify_exp": True,
            "verify_aud": True,
            "verify_iss": True,
        },
    )
    return claims


# ---------------------------------------------------------------------------
# Flask decorator
# ---------------------------------------------------------------------------
def require_bearer_token(f):
    """
    Flask route decorator: validates the Authorization Bearer token.

    Returns 401 with a JSON body if validation fails.
    Sets `request.token_claims` on success for downstream use.
    """

    @wraps(f)
    def decorated(*args, **kwargs):
        auth_header = request.headers.get("Authorization", "")

        if not auth_header.startswith("Bearer "):
            logger.warning("Missing or malformed Authorization header")
            return jsonify({
                "error": "unauthorized",
                "message": "Authorization header must be 'Bearer <token>'"
            }), 401

        token = auth_header[7:]  # strip "Bearer "

        try:
            claims = _validate_token(token)
            request.token_claims = claims  # type: ignore[attr-defined]
        except jwt.ExpiredSignatureError:
            logger.warning("Bearer token has expired")
            return jsonify({"error": "token_expired",
                            "message": "Bearer token has expired"}), 401
        except jwt.InvalidAudienceError:
            logger.warning("Bearer token audience mismatch")
            return jsonify({"error": "invalid_audience",
                            "message": "Token audience does not match this application"}), 401
        except jwt.InvalidIssuerError:
            logger.warning("Bearer token issuer mismatch")
            return jsonify({"error": "invalid_issuer",
                            "message": "Token issuer is not trusted"}), 401
        except jwt.InvalidTokenError as exc:
            logger.warning(f"Bearer token validation failed: {exc}")
            return jsonify({"error": "invalid_token",
                            "message": str(exc)}), 401
        except requests.RequestException as exc:
            logger.error(f"Failed to fetch JWKS for token validation: {exc}")
            return jsonify({"error": "auth_service_error",
                            "message": "Unable to validate token at this time"}), 503

        return f(*args, **kwargs)

    return decorated

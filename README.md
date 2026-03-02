# Azure AI Foundry Agent + SharePoint with OBO Identity Passthrough

When an Azure AI Foundry agent needs to access SharePoint, it must do so **as the signed-in user** — not as an app — so that SharePoint's per-document permissions are enforced. This is the On-Behalf-Of (OBO) identity challenge: the user's token must flow from the client, through Foundry, and into SharePoint without being replaced by a service identity. Getting this chain right is the difference between a secure, permission-respecting agent and one that leaks documents users shouldn't see.

This repo demonstrates **three patterns** for solving this, from simplest to most flexible.

## Patterns at a Glance

| | Pattern 1 | Pattern 2 | Pattern 3 |
|---|---|---|---|
| **Folder** | `01-foundry-sharepoint-tool/` | `02-m365-retrieval-api/` | `03-declarative-agent-manifest/` |
| **Approach** | Foundry native SharePoint tool | M365 Copilot Retrieval API + Foundry | Declarative Agent manifest + Foundry plugin |
| **OBO mechanism** | Identity passthrough (SDK-managed) | Delegated Graph token (you manage) | M365 Copilot handles it |
| **Code required** | Python (~50 lines) | Python (~80 lines) | JSON manifests + stub endpoint |
| **M365 Copilot licence** | Required | Required (or pay-as-you-go) | Required |
| **Runs in M365 Copilot UI** | No (standalone app) | No (standalone app) | Yes (Teams / M365 Chat) |
| **SharePoint permission enforcement** | Automatic via Foundry | Automatic via Graph API | Automatic via M365 Copilot |
| **Custom orchestration** | Limited (agent instructions only) | Full control | Moderate (declarative + plugin) |
| **Best for** | Quick proof-of-concept | Custom apps needing full control | Enterprise deployment in M365 |

## Prerequisites

All patterns require:

- **Azure subscription** with an [Azure AI Foundry](https://ai.azure.com) project
- **Microsoft 365 Copilot licence** on the user account (or pay-as-you-go for Pattern 2)
- **Azure AD app registration** with appropriate permissions
- **Python 3.9+** (for Patterns 1 and 2)

## Quick Start

### Pattern 1: Foundry Native SharePoint Tool

The simplest approach — Foundry handles the OBO token exchange automatically.

```bash
cd 01-foundry-sharepoint-tool
pip install -r requirements.txt
cp .env.example .env
# Edit .env with your values
python main.py
```

→ [Pattern 1 README](01-foundry-sharepoint-tool/README.md)

### Pattern 2: M365 Copilot Retrieval API

Full control over the retrieval and synthesis pipeline. You acquire the user's token, call the Retrieval API, then feed results to Foundry.

```bash
cd 02-m365-retrieval-api
pip install -r requirements.txt
cp .env.example .env
# Edit .env with your values
python main.py
```

→ [Pattern 2 README](02-m365-retrieval-api/README.md)

### Pattern 3: Declarative Agent Manifest

No Python required for the agent itself — just JSON manifests deployed to M365 Copilot, with an optional Foundry-backed API plugin.

```bash
cd 03-declarative-agent-manifest
# Deploy declarative-agent.json via Teams Developer Portal or Teams Toolkit
# Optionally deploy stub-endpoint/ to Azure App Service
```

→ [Pattern 3 README](03-declarative-agent-manifest/README.md)

## How OBO Works in Each Pattern

```
Pattern 1 (Foundry native):
  User → InteractiveBrowserCredential → Foundry SDK → [identity passthrough] → SharePoint

Pattern 2 (Retrieval API):
  User → MSAL interactive login → Graph delegated token → Retrieval API → SharePoint
  Retrieved content → Foundry agent (service identity) → Response

Pattern 3 (Declarative Agent):
  User → M365 Copilot (manages OBO) → SharePoint grounding → Plugin call → Foundry endpoint
```

## Official Documentation

- [Azure AI Foundry Agents overview](https://learn.microsoft.com/azure/ai-services/agents/overview)
- [SharePoint tool for Foundry Agents](https://learn.microsoft.com/azure/ai-services/agents/how-to/tools/sharepoint)
- [M365 Copilot Retrieval API (beta)](https://learn.microsoft.com/graph/api/resources/copilot-retrieval?view=graph-rest-beta)
- [Declarative agents for M365 Copilot](https://learn.microsoft.com/microsoft-365-copilot/extensibility/overview-declarative-agent)
- [On-Behalf-Of flow (Microsoft identity platform)](https://learn.microsoft.com/entra/identity-platform/v2-oauth2-on-behalf-of-flow)

## ⚠️ Important Notes

- **This is a reference sample**, not production code. Error handling, logging, and retry logic are minimal.
- **M365 Copilot licence** is required for SharePoint grounding in all patterns. Without it, the SharePoint tool and Retrieval API will not return results.
- **Identity passthrough** means the agent can only see what the user can see. This is a feature, not a bug.
- **Pattern 1 cannot be published to M365 Copilot** directly — the identity chain breaks when Foundry agents are surfaced through Teams. Use Pattern 3 for that scenario.

## Contributing

This is a sample repo for reference. Feel free to fork and adapt to your needs.

## License

MIT

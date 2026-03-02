"""
Stub Flask endpoint — replace with your actual Foundry agent call.
Deploy to Azure App Service, Container App, or Azure Functions.
"""
from flask import Flask, request, jsonify
from azure.ai.projects import AIProjectClient
from azure.identity import DefaultAzureCredential
import os

app = Flask(__name__)


@app.route("/analyze", methods=["POST"])
def analyze():
    data = request.json
    query = data.get("query", "")
    context = data.get("context", "")

    # Call your Foundry agent here
    client = AIProjectClient(
        endpoint=os.environ["FOUNDRY_PROJECT_ENDPOINT"],
        credential=DefaultAzureCredential(),
    )

    thread = client.agents.create_thread()
    client.agents.create_message(
        thread_id=thread.id,
        role="user",
        content=f"Query: {query}\n\nContext from SharePoint:\n{context}",
    )
    run = client.agents.create_and_process_run(
        thread_id=thread.id,
        agent_id=os.environ["FOUNDRY_AGENT_ID"],
    )
    messages = client.agents.list_messages(thread_id=thread.id)
    for msg in messages:
        if msg.role == "assistant":
            return jsonify({"analysis": msg.content[0].text.value, "confidence": 0.9})

    return jsonify({"analysis": "No response generated", "confidence": 0.0})


if __name__ == "__main__":
    app.run(port=8000)

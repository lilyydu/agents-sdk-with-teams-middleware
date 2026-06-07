"""Phase 3 — curated, typed invoke decorators.

The day-to-day developer experience. Each decorator selects on the right
``activity.name`` (and where useful, a sub-key like ``commandId``) and
the response body is wrapped into a Teams ``InvokeResponse`` automatically.
"""

from microsoft_agents.hosting.core import AgentApplication
from microsoft_agents.hosting.teams import install_teams, TeamsHandlers

app = AgentApplication()
install_teams(app)
teams = TeamsHandlers(app)


# ── Message extension search ──
@teams.message_extension_query("search")
async def search(ctx, query):
    # TODO: build a real MessagingExtensionResponse
    return {
        "composeExtension": {
            "type": "result",
            "attachmentLayout": "list",
            "attachments": [],
        }
    }


# ── Task module open (modal) ──
@teams.task_module_fetch("createTicket")
async def fetch_create_ticket(ctx, request):
    return {
        "task": {
            "type": "continue",
            "value": {"title": "Create ticket", "card": {}},
        }
    }


# ── Task module submit ──
@teams.task_module_submit("createTicket")
async def submit_create_ticket(ctx, request):
    return {"task": {"type": "message", "value": "Ticket created."}}


# ── Adaptive Card Action.Execute ──
@teams.adaptive_card_action_execute("approve")
async def approve(ctx, data):
    return {"type": "AdaptiveCard", "version": "1.5", "body": []}


# ── Regular chat still works alongside invokes ──
@app.message("status")
async def status(ctx, state):
    await ctx.send_activity("All good.")

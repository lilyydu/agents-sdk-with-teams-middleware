"""
Middleware that matches Teams turns to a teams.py App.

Lifecycle of a turn:

    1. Non-Teams channel → pass through to AgentApplication.
    2. Teams turn with no matching teams.py route → pass through too.
    3. Teams turn with a match → ensure teams.py is initialized, expose the
       Agents SDK TurnContext via a ContextVar, hand the activity to
       teams.py's activity_processor, then propagate any InvokeResponse
       back through the Agents SDK send pipeline so the HTTP layer can
       write the synchronous body.

The middleware never calls ``logic()`` after teams.py has handled a turn —
doing so would invoke AgentApplication's handlers a second time for the
same activity and, for invokes, would emit a duplicate response.
"""

from __future__ import annotations

from typing import Awaitable, Callable

from microsoft_agents.activity import Activity, ActivityTypes, InvokeResponse
from microsoft_agents.hosting.core.middleware_set import Middleware
from microsoft_agents.hosting.core.turn_context import TurnContext

from microsoft_teams.api import ActivityTypeAdapter
from microsoft_teams.apps import App
from microsoft_teams.apps.events.types import ActivityEvent

from ._context import _agent_turn_context
from ._token import _TeamsSDKToken

TEAMS_CHANNEL_ID = "msteams"


class TeamsSDKMiddleware(Middleware):

    def __init__(self, teams_app: App) -> None:
        self._teams_app = teams_app

    async def on_turn(
        self,
        context: TurnContext,
        logic: Callable[[], Awaitable],
    ) -> None:
        if context.activity.channel_id != TEAMS_CHANNEL_ID:
            await logic()
            return

        core_activity = self._translate_inbound(context.activity)

        # Idempotent — App tracks _initialized internally. Run on every Teams
        # turn so AgentApplication handlers can safely call into TEAMS_APP
        # (e.g. ``TEAMS_APP.send`` for proactive sends) even when no teams.py
        # route matched this turn.
        await self._teams_app.initialize()

        if not self._teams_app.router.select_handlers(core_activity):
            # No teams.py route matches; let AgentApplication try its handlers.
            await logic()
            return

        event = ActivityEvent(
            body=core_activity,
            token=_TeamsSDKToken.from_activity(context.activity),
        )

        ctx_token = _agent_turn_context.set(context)
        try:
            invoke_response = await self._teams_app.activity_processor.process_activity(
                plugins=[], event=event
            )
        finally:
            _agent_turn_context.reset(ctx_token)

        if context.activity.type == ActivityTypes.invoke:
            await self._propagate_invoke_response(context, invoke_response)

    @staticmethod
    def _translate_inbound(activity: Activity):
        """Agents SDK Activity → teams.py typed Activity subclass via JSON.

        Both sides are Pydantic but distinct model hierarchies. Going through
        JSON keeps the bridge resilient to additive schema drift.
        """
        payload = activity.model_dump_json(by_alias=True, exclude_none=True)
        return ActivityTypeAdapter.validate_json(payload)

    @staticmethod
    async def _propagate_invoke_response(
        context: TurnContext, invoke_response: object
    ) -> None:
        """Emit teams.py's InvokeResponse through the Agents SDK send pipeline.

        Wraps the response in an outbound activity of type ``invoke_response``;
        ChannelServiceAdapter stashes it under INVOKE_RESPONSE_KEY so the HTTP
        layer writes it as the synchronous response body.

        Accepts both ``InvokeResponse`` instances and the dict shape
        (``{"status": int, "body": ...}``) that teams.py's ``is_invoke_response``
        treats as already-an-InvokeResponse. Pydantic bodies are dumped to
        plain dicts so stdlib ``json.dumps`` downstream can encode them.
        """
        if invoke_response is None:
            return

        if isinstance(invoke_response, dict):
            status = invoke_response.get("status")
            body = invoke_response.get("body")
        else:
            status = getattr(invoke_response, "status", None)
            body = getattr(invoke_response, "body", None)

        if status is None:
            return

        if hasattr(body, "model_dump"):
            body = body.model_dump(mode="json", by_alias=True, exclude_none=True)

        await context.send_activity(
            Activity(
                type=ActivityTypes.invoke_response,
                value=InvokeResponse(status=status, body=body),
            )
        )

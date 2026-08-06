"""
Middleware that matches Teams turns to a teams.py App.

Lifecycle of a turn:

    1. Non-Teams channel → pass through to AgentApplication.
    2. Teams turn the caller's ``should_bypass_teams`` predicate claims → pass
       through too, even when teams.py has a matching route.
    3. Teams turn with no matching teams.py route → pass through too.
    4. Teams turn with a match → ensure teams.py is initialized, expose the
       Agents SDK TurnContext via a ContextVar, hand the activity to
       teams.py's activity_processor, then propagate any InvokeResponse
       back through the Agents SDK send pipeline so the HTTP layer can
       write the synchronous body.

The middleware never calls ``logic()`` after teams.py has handled a turn —
doing so would invoke AgentApplication's handlers a second time for the
same activity and, for invokes, would emit a duplicate response.
"""

from __future__ import annotations

from typing import Awaitable, Callable, Optional

from microsoft_agents.activity import Activity, ActivityTypes, InvokeResponse
from microsoft_agents.hosting.core.middleware_set import Middleware
from microsoft_agents.hosting.core.turn_context import TurnContext

from microsoft_teams.api import ActivityTypeAdapter
from microsoft_teams.apps import App
from microsoft_teams.apps.events.types import ActivityEvent

from ._context import _agent_sdk_turn_context
from ._token import _TeamsSDKToken

TEAMS_CHANNEL_ID = "msteams"


def is_teams_channel(activity) -> bool:
    """True for Teams turns, including sub-channels like ``msteams:COPILOT``."""
    channel_id = activity.channel_id
    if not channel_id:
        return False

    channel = getattr(channel_id, "channel", None)
    if channel is None:
        channel = str(channel_id).split(":", 1)[0]
    return channel == TEAMS_CHANNEL_ID


class TeamsSDKMiddleware(Middleware):

    def __init__(
        self,
        teams_app: App,
        should_bypass_teams: Optional[Callable[[TurnContext], bool]] = None,
    ) -> None:
        """
        Args:
            teams_app: The teams.py ``App`` that owns matching Teams turns.
            should_bypass_teams: Optional predicate evaluated only for
                Teams-channel turns. Return ``True`` to force the turn to fall
                through to ``AgentApplication`` even when teams.py has a
                matching route.
        """
        self._teams_app = teams_app
        self._should_bypass_teams = should_bypass_teams

    async def on_turn(
        self,
        context: TurnContext,
        logic: Callable[[TurnContext], Awaitable],
    ) -> None:
        if not is_teams_channel(context.activity):
            await logic(context)
            return

        # Idempotent — App tracks _initialized internally. Run on every Teams
        # turn so AgentApplication handlers can safely call into TEAMS_APP
        # (e.g. ``TEAMS_APP.send`` for proactive sends) even when no teams.py
        # route matched this turn.
        await self._teams_app.initialize()

        if self._should_bypass_teams is not None and self._should_bypass_teams(context):
            await logic(context)
            return

        core_activity = self._translate_inbound(context.activity)

        if not self._teams_app.router.select_handlers(core_activity):
            # No teams.py route matches; let AgentApplication try its handlers.
            await logic(context)
            return

        event = ActivityEvent(
            body=core_activity,
            token=_TeamsSDKToken.from_activity(context.activity),
        )

        ctx_token = _agent_sdk_turn_context.set(context)
        try:
            invoke_response = await self._teams_app.activity_processor.process_activity(
                plugins=[], event=event
            )
        finally:
            _agent_sdk_turn_context.reset(ctx_token)

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

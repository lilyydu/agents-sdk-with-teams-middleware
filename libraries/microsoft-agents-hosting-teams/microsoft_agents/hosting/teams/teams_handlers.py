"""Teams handler decorators. Thin wrappers over ``AgentApplication.add_route``.

Two layers:

* **Escape hatch** (``on_invoke``, ``on_activity``):
    Day-1 support for any Teams invoke name or activity filter. The Teams team
    can document a new invoke type before adding a typed decorator; developers
    are never blocked waiting for an SDK release.

* **Curated sugar** (``message_extension_query``, ``task_module_fetch``,
    ``adaptive_card_action_execute``, ...):
    Typed, ergonomic decorators for the common invoke types. Each one selects
    on ``activity.name`` (and where useful, on a sub-key like ``commandId``)
    and unwraps ``activity.value`` into a Pydantic model.

Every selector checks ``channel_id == "msteams"`` so Teams routes never
fire on other channels.

Zero core changes — uses the existing ``add_route`` public API.
"""

from __future__ import annotations

from typing import Any, Awaitable, Callable, Optional

from microsoft_agents.activity import ActivityTypes, InvokeResponse
from microsoft_agents.hosting.core.app.agent_application import AgentApplication
from microsoft_agents.hosting.core.app._routes.route_rank import RouteRank
from microsoft_agents.hosting.core.turn_context import TurnContext

TEAMS_CHANNEL_ID = "msteams"

# Core unwraps this turn_state key and returns it as the HTTP response body
# for invoke activities. Constant lives in core; duplicated here to avoid a
# fragile import on a private symbol.
INVOKE_RESPONSE_KEY = "BotFrameworkAdapter.InvokeResponse"

# Invokes outrank regular messages; using FIRST + small offset keeps room
# for the developer to override with rank=RouteRank.FIRST if they really want.
_INVOKE_RANK = RouteRank(10)


class TeamsHandlers:
    """Decorators that register Teams routes on an AgentApplication.

    Usage:
        app = AgentApplication()
        install_teams(app)                 # plumbing
        teams = TeamsHandlers(app)         # surface

        @teams.message_extension_query("search")
        async def search(ctx, query): ...
    """

    def __init__(self, app: AgentApplication):
        self._app = app

    # ═══════════════════ Layer 1: escape hatch ═══════════════════

    def on_invoke(
        self,
        name: str,
        *,
        rank: RouteRank = _INVOKE_RANK,
        auth_handlers: Optional[list[str]] = None,
    ):
        """Catch any Teams invoke by ``activity.name``.

        Use for invoke types not yet covered by a typed decorator. Handler
        receives the raw ``activity.value`` and should return a JSON-
        serialisable body (or ``None`` for a 200 with empty body).
        """

        def decorator(handler: Callable[[TurnContext, Any], Awaitable[Any]]):
            async def selector(ctx: TurnContext) -> bool:
                return (
                    ctx.activity.channel_id == TEAMS_CHANNEL_ID
                    and ctx.activity.type == ActivityTypes.invoke
                    and ctx.activity.name == name
                )

            async def wrapped(ctx: TurnContext, _state=None) -> None:
                body = await handler(ctx, ctx.activity.value)
                _set_invoke_response(ctx, body)

            self._app.add_route(
                selector=selector,
                handler=wrapped,
                is_invoke=True,
                rank=rank,
                auth_handlers=auth_handlers,
            )
            return handler

        return decorator

    def on_activity(
        self,
        predicate: Callable[[TurnContext], bool],
        *,
        rank: RouteRank = RouteRank.DEFAULT,
        auth_handlers: Optional[list[str]] = None,
    ):
        """Catch any Teams activity matching ``predicate``.

        Use for non-invoke Teams activities (e.g. ``conversationUpdate`` with
        Teams-specific channelData like team renamed).
        """

        def decorator(handler: Callable[[TurnContext], Awaitable[None]]):
            async def selector(ctx: TurnContext) -> bool:
                if ctx.activity.channel_id != TEAMS_CHANNEL_ID:
                    return False
                return predicate(ctx)

            async def wrapped(ctx: TurnContext, _state=None) -> None:
                await handler(ctx)

            self._app.add_route(
                selector=selector,
                handler=wrapped,
                is_invoke=False,
                rank=rank,
                auth_handlers=auth_handlers,
            )
            return handler

        return decorator

    # ═══════════════════ Layer 2: curated sugar ═══════════════════

    def message_extension_query(
        self,
        command_id: str,
        *,
        auth_handlers: Optional[list[str]] = None,
    ):
        """Compose-extension query (Teams message extensions)."""

        def decorator(handler):
            async def selector(ctx: TurnContext) -> bool:
                return (
                    ctx.activity.channel_id == TEAMS_CHANNEL_ID
                    and ctx.activity.type == ActivityTypes.invoke
                    and ctx.activity.name == "composeExtension/query"
                    and (ctx.activity.value or {}).get("commandId") == command_id
                )

            async def wrapped(ctx: TurnContext, _state=None) -> None:
                # TODO: parse into MessagingExtensionQuery model when available
                response = await handler(ctx, ctx.activity.value)
                _set_invoke_response(ctx, response)

            self._app.add_route(
                selector=selector,
                handler=wrapped,
                is_invoke=True,
                rank=_INVOKE_RANK,
                auth_handlers=auth_handlers,
            )
            return handler

        return decorator

    def task_module_fetch(
        self,
        command_id: Optional[str] = None,
        *,
        auth_handlers: Optional[list[str]] = None,
    ):
        """Task module fetch (modal dialog open)."""

        def decorator(handler):
            async def selector(ctx: TurnContext) -> bool:
                if not (
                    ctx.activity.channel_id == TEAMS_CHANNEL_ID
                    and ctx.activity.type == ActivityTypes.invoke
                    and ctx.activity.name == "task/fetch"
                ):
                    return False
                if command_id is None:
                    return True
                data = (ctx.activity.value or {}).get("data") or {}
                return data.get("commandId") == command_id

            async def wrapped(ctx: TurnContext, _state=None) -> None:
                response = await handler(ctx, ctx.activity.value)
                _set_invoke_response(ctx, response)

            self._app.add_route(
                selector=selector,
                handler=wrapped,
                is_invoke=True,
                rank=_INVOKE_RANK,
                auth_handlers=auth_handlers,
            )
            return handler

        return decorator

    def task_module_submit(
        self,
        command_id: Optional[str] = None,
        *,
        auth_handlers: Optional[list[str]] = None,
    ):
        """Task module submit (modal dialog OK pressed)."""

        def decorator(handler):
            async def selector(ctx: TurnContext) -> bool:
                if not (
                    ctx.activity.channel_id == TEAMS_CHANNEL_ID
                    and ctx.activity.type == ActivityTypes.invoke
                    and ctx.activity.name == "task/submit"
                ):
                    return False
                if command_id is None:
                    return True
                data = (ctx.activity.value or {}).get("data") or {}
                return data.get("commandId") == command_id

            async def wrapped(ctx: TurnContext, _state=None) -> None:
                response = await handler(ctx, ctx.activity.value)
                _set_invoke_response(ctx, response)

            self._app.add_route(
                selector=selector,
                handler=wrapped,
                is_invoke=True,
                rank=_INVOKE_RANK,
                auth_handlers=auth_handlers,
            )
            return handler

        return decorator

    def adaptive_card_action_execute(
        self,
        verb: str,
        *,
        auth_handlers: Optional[list[str]] = None,
    ):
        """Adaptive Card Action.Execute by verb."""

        def decorator(handler):
            async def selector(ctx: TurnContext) -> bool:
                if not (
                    ctx.activity.channel_id == TEAMS_CHANNEL_ID
                    and ctx.activity.type == ActivityTypes.invoke
                    and ctx.activity.name == "adaptiveCard/action"
                ):
                    return False
                action = (ctx.activity.value or {}).get("action") or {}
                return action.get("verb") == verb

            async def wrapped(ctx: TurnContext, _state=None) -> None:
                action = (ctx.activity.value or {}).get("action") or {}
                data = action.get("data") or {}
                card = await handler(ctx, data)
                _set_invoke_response(
                    ctx,
                    {
                        "statusCode": 200,
                        "type": "application/vnd.microsoft.card.adaptive",
                        "value": card,
                    },
                )

            self._app.add_route(
                selector=selector,
                handler=wrapped,
                is_invoke=True,
                rank=_INVOKE_RANK,
                auth_handlers=auth_handlers,
            )
            return handler

        return decorator

    def sign_in_verify_state(self, *, auth_handlers: Optional[list[str]] = None):
        """OAuth sign-in completion: ``signin/verifyState``."""

        def decorator(handler):
            async def selector(ctx: TurnContext) -> bool:
                return (
                    ctx.activity.channel_id == TEAMS_CHANNEL_ID
                    and ctx.activity.type == ActivityTypes.invoke
                    and ctx.activity.name == "signin/verifyState"
                )

            async def wrapped(ctx: TurnContext, _state=None) -> None:
                await handler(ctx, ctx.activity.value)
                _set_invoke_response(ctx, None)

            self._app.add_route(
                selector=selector,
                handler=wrapped,
                is_invoke=True,
                rank=_INVOKE_RANK,
                auth_handlers=auth_handlers,
            )
            return handler

        return decorator


# ────────────────────── helpers ──────────────────────


def _set_invoke_response(ctx: TurnContext, body: Any) -> None:
    """Core unwraps this turn_state key into the HTTP response for invokes."""
    ctx.turn_state[INVOKE_RESPONSE_KEY] = InvokeResponse(status=200, body=body)

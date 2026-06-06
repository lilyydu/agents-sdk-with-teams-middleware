"""TeamsDeepExtension — rich, opt-in handler surface for Teams.

Two layers of decorators:

  * Layer 1 — escape hatch. Always works, NEVER lags Teams releases.
      @teams.on_invoke("composeExtension/newThing")
      @teams.on_activity(lambda a: ...)

  * Layer 2 — curated typed sugar. Built on top of Layer 1. Each typed
      wrapper is one short function that matches by invoke name and validates
      `activity.value` into a typed Pydantic payload.

`can_handle` returns True for Teams traffic the dev registered a route for.
For unregistered turns (e.g. ordinary Teams text messages with no
`@teams.on_message` handler), it returns False so the generic cross-platform
`@app.on_message` handler can run.
"""

from __future__ import annotations

from http import HTTPStatus
from typing import Any, Callable, Optional, TYPE_CHECKING

from microsoft_agents.activity import Activity, ActivityTypes, InvokeResponse

from microsoft_agents.hosting.core.extension import DeepExtension
from microsoft_agents.hosting.core.turn_context import TurnContext

from .teams_context import TeamsContext

if TYPE_CHECKING:
    from microsoft_agents.hosting.core.app.agent_application import AgentApplication


CHANNEL_ID = "msteams"

ActivityFilter = Callable[[Activity], bool]
Handler = Callable[..., Any]


class _Route:
    """Internal route record."""

    def __init__(
        self,
        filter: ActivityFilter,
        handler: Handler,
        *,
        payload_type: Optional[type] = None,
    ) -> None:
        self.filter = filter
        self.handler = handler
        self.payload_type = payload_type


class TeamsDeepExtension(DeepExtension):
    name = "teams-deep-extension"
    version = "0.1.0"
    requires_core = ">=2.0,<3.0"

    def __init__(self) -> None:
        self._routes: list[_Route] = []
        self._host: Optional["AgentApplication"] = None

        # Sub-surfaces — populated below for ergonomic grouping.
        self.message_extension = _MessageExtensionSurface(self)
        self.task_module = _TaskModuleSurface(self)
        self.meeting = _MeetingSurface(self)

    def initialize(self, host: "AgentApplication") -> None:
        self._host = host

    # ------------------------------------------------------------ contract
    def owns_channel(self, channel_id: str) -> bool:
        return channel_id == CHANNEL_ID

    def can_handle(self, ctx: TurnContext) -> bool:
        return any(r.filter(ctx.activity) for r in self._routes)

    async def handle(self, ctx: TurnContext) -> None:
        tctx = TeamsContext.from_turn_context(ctx)
        for route in self._routes:
            if not route.filter(ctx.activity):
                continue
            payload = self._coerce_payload(ctx.activity, route.payload_type)
            result = await route.handler(tctx, payload) if payload is not None \
                else await route.handler(tctx)
            if ctx.activity.type == ActivityTypes.invoke:
                self._set_invoke_response(ctx, result)
            return

    # ======================================================== Layer 1 hatch

    def on_invoke(
        self,
        name: str,
        *,
        payload_type: Optional[type] = None,
    ) -> Callable[[Handler], Handler]:
        """Register a handler for any invoke `activity.name`.

        Always available — no extension release needed when Teams ships a
        new invoke type.
        """
        def filter_fn(a: Activity) -> bool:
            return a.type == ActivityTypes.invoke and a.name == name

        def decorator(fn: Handler) -> Handler:
            self._routes.append(_Route(filter_fn, fn, payload_type=payload_type))
            return fn

        return decorator

    def on_activity(
        self,
        filter: ActivityFilter,
        *,
        payload_type: Optional[type] = None,
    ) -> Callable[[Handler], Handler]:
        """Register a handler for any activity matching `filter(activity)`."""

        def decorator(fn: Handler) -> Handler:
            self._routes.append(_Route(filter, fn, payload_type=payload_type))
            return fn

        return decorator

    # ======================================================== Layer 2 sugar

    def on_message(self) -> Callable[[Handler], Handler]:
        """Phase-4 override: Teams text messages handled by this extension."""
        return self.on_activity(lambda a: a.type == ActivityTypes.message)

    def on_message_reaction(self) -> Callable[[Handler], Handler]:
        return self.on_activity(lambda a: a.type == ActivityTypes.message_reaction)

    def on_read_receipt(self) -> Callable[[Handler], Handler]:
        return self.on_activity(
            lambda a: a.type == ActivityTypes.event and a.name == "application/vnd.microsoft.readReceipt"
        )

    # ... add typed sugar as the team chooses; each is a one-liner over the
    # escape hatch. NEW Teams things never block on adding these.

    # =========================================================== internals

    def _coerce_payload(self, activity: Activity, payload_type: Optional[type]) -> Any:
        if payload_type is None:
            return activity.value
        if hasattr(payload_type, "model_validate"):
            return payload_type.model_validate(activity.value or {})
        return activity.value

    def _set_invoke_response(self, ctx: TurnContext, result: Any) -> None:
        body: Any = None
        if result is not None:
            body = result.model_dump(mode="json", by_alias=True, exclude_none=True) \
                if hasattr(result, "model_dump") else result
        ctx.turn_state["invoke_response_value"] = InvokeResponse(
            status=int(HTTPStatus.OK), body=body
        )


# ---------------------------------------------------------------------------
# Sub-surfaces — ergonomic grouping; same escape-hatch underneath.
# ---------------------------------------------------------------------------

class _MessageExtensionSurface:
    def __init__(self, parent: TeamsDeepExtension) -> None:
        self._parent = parent

    def on_query(self, command_id: Optional[str] = None):
        from microsoft_teams.api.models import MessagingExtensionQuery

        def filter_fn(a: Activity) -> bool:
            if a.type != ActivityTypes.invoke or a.name != "composeExtension/query":
                return False
            if command_id is None:
                return True
            v = a.value or {}
            cmd = v.get("commandId") if isinstance(v, dict) else getattr(v, "commandId", None)
            return cmd == command_id

        return self._parent.on_activity(filter_fn, payload_type=MessagingExtensionQuery)

    def on_submit_action(self, command_id: Optional[str] = None):
        from microsoft_teams.api.models import MessagingExtensionAction

        def filter_fn(a: Activity) -> bool:
            if a.type != ActivityTypes.invoke or a.name != "composeExtension/submitAction":
                return False
            if command_id is None:
                return True
            v = a.value or {}
            cmd = v.get("commandId") if isinstance(v, dict) else getattr(v, "commandId", None)
            return cmd == command_id

        return self._parent.on_activity(filter_fn, payload_type=MessagingExtensionAction)


class _TaskModuleSurface:
    def __init__(self, parent: TeamsDeepExtension) -> None:
        self._parent = parent

    def on_fetch(self):
        from microsoft_teams.api.models import TaskModuleRequest

        return self._parent.on_invoke("task/fetch", payload_type=TaskModuleRequest)

    def on_submit(self):
        from microsoft_teams.api.models import TaskModuleRequest

        return self._parent.on_invoke("task/submit", payload_type=TaskModuleRequest)


class _MeetingSurface:
    def __init__(self, parent: TeamsDeepExtension) -> None:
        self._parent = parent

    def on_start(self):
        return self._parent.on_activity(
            lambda a: a.type == ActivityTypes.event
            and a.name == "application/vnd.microsoft.meetingStart"
        )

    def on_end(self):
        return self._parent.on_activity(
            lambda a: a.type == ActivityTypes.event
            and a.name == "application/vnd.microsoft.meetingEnd"
        )

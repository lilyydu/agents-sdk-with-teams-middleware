"""TeamsContext — rich context handed to TeamsDeepExtension handlers.

Built from a TurnContext via a bridge (not subclassing). The bridge:

  * exposes a Teams-flavored API surface (`send`, `stream`, `graph`, `signin`,
    `mention`, `meeting`, `cards`, `read_receipt`, ...).
  * funnels ALL outbound traffic back through the standard TurnContext, so
    `TeamsChannelAdapter.on_activity_sending` runs on every send — no drift.

Important: there is exactly one egress (TurnContext.send_activity), so a
handler that holds both `tctx.send(...)` and `ctx.send_activity(...)` gets
identical Teams transforms either way.
"""

from __future__ import annotations

from typing import Any, Optional

from microsoft_agents.activity import Activity

from microsoft_agents.hosting.core.turn_context import TurnContext


class TeamsContext:
    """Bridge wrapper. Holds the original TurnContext and exposes Teams API."""

    def __init__(self, turn_context: TurnContext) -> None:
        self._ctx = turn_context

    # ----- bridge factory -----
    @classmethod
    def from_turn_context(cls, ctx: TurnContext) -> "TeamsContext":
        return cls(ctx)

    # ----- pass-throughs -----
    @property
    def activity(self) -> Activity:
        return self._ctx.activity

    @property
    def turn_state(self) -> dict:
        return self._ctx.turn_state

    @property
    def turn_context(self) -> TurnContext:
        return self._ctx

    # ----- Teams API surface (delegates to TurnContext for egress) -----

    async def send(self, text_or_activity) -> Any:
        """Send a message. Single egress → on_activity_sending always runs."""
        return await self._ctx.send_activity(text_or_activity)

    async def stream(self, text: str) -> None:
        """Stream a partial response (Teams streaming protocol).

        Implementation builds intermediate `typing`/`message` activities and
        sends them via self._ctx.send_activity so transforms still apply.
        """
        raise NotImplementedError

    @property
    def graph(self):
        """Microsoft Graph client scoped to the signed-in user.

        Resolved lazily from `turn_state["teams.graph"]` set by
        TeamsChannelAdapter when SSO is configured.
        """
        return self._ctx.turn_state.get("teams.graph")

    async def signin(self, connection_name: str) -> Optional[str]:
        """Request user sign-in via core's Authorization. Returns token if any."""
        # Pseudocode: auth = self._ctx.turn_state["app"].auth
        # return await auth.sign_in(self._ctx, connection_name)
        raise NotImplementedError

    def mention(self, user_id: str, name: str) -> dict:
        """Build a Teams <at> mention; TeamsChannelAdapter wires it on egress."""
        return {"type": "mention", "mentioned": {"id": user_id, "name": name}, "text": f"<at>{name}</at>"}

    # ----- response helpers -----

    def respond_to_invoke(self, body: Any) -> None:
        """Set the invoke response body for this turn (alternative to returning
        a value from the handler).
        """
        from http import HTTPStatus
        from microsoft_agents.activity import InvokeResponse

        self._ctx.turn_state["invoke_response_value"] = InvokeResponse(
            status=int(HTTPStatus.OK),
            body=body.model_dump(mode="json", by_alias=True, exclude_none=True)
            if hasattr(body, "model_dump")
            else body,
        )

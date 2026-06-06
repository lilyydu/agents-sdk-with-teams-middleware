"""TeamsChannelAdapter — channel-scoped plumbing for Microsoft Teams.

Registers with `AgentApplication.extensions` at `initialize()` time. After
registration, core calls back into:

  * on_context_created(ctx)    — swap connector client, attach Teams/Graph
                                  clients, parse channelData.
  * on_activity_sending(ctx, a)— transform outbound activities (mentions,
                                  channelData, card variants, streaming).

This piece runs for EVERY Teams turn, including cross-platform handlers like
`@app.on_message`. Devs who never write a Teams-specific handler still get
Teams-correct plumbing once they `app.use(TeamsChannelAdapter(...))`.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Optional

from microsoft_agents.activity import Activity

from microsoft_agents.hosting.core.extension import ChannelAdapter
from microsoft_agents.hosting.core.turn_context import TurnContext
from microsoft_agents.hosting.core.connector.teams import TeamsConnectorClient

if TYPE_CHECKING:
    from microsoft_agents.hosting.core.app.agent_application import AgentApplication


CHANNEL_ID = "msteams"


class TeamsChannelAdapter(ChannelAdapter):
    name = "teams-channel-adapter"
    version = "0.1.0"
    requires_core = ">=2.0,<3.0"

    def __init__(
        self,
        *,
        app_id: str,
        app_password: Optional[str] = None,
    ) -> None:
        self._app_id = app_id
        self._app_password = app_password
        self._host: Optional["AgentApplication"] = None

    def initialize(self, host: "AgentApplication") -> None:
        """Optional lifecycle hook called by AgentApplication.use(...)."""
        self._host = host

    def owns_channel(self, channel_id: str) -> bool:
        return channel_id == CHANNEL_ID

    # ------------------------------------------------------------------ in
    async def on_context_created(self, ctx: TurnContext) -> None:
        """Swap to a TeamsConnectorClient and enrich the context."""
        token = await self._acquire_teams_token(ctx)
        teams_client = TeamsConnectorClient(
            endpoint=ctx.activity.service_url,
            token=token,
        )
        # Single egress for Teams traffic — owned by the Teams team.
        ctx.turn_state["ConnectorClient"] = teams_client

        # Stash typed Teams helpers for any handler (cross-platform or deep).
        ctx.turn_state["teams.connector_client"] = teams_client
        ctx.turn_state["teams.channel_data"] = self._parse_channel_data(ctx.activity)
        # Optional: attach a Graph client if the user has SSO configured.
        # ctx.turn_state["teams.graph"] = await self._maybe_graph_client(ctx)

    # ----------------------------------------------------------------- out
    async def on_activity_sending(
        self, ctx: TurnContext, activity: Activity
    ) -> Activity:
        """Apply Teams-specific transforms to outbound activities."""
        activity = self._inject_channel_data(ctx, activity)
        activity = self._inject_mentions(activity)
        activity = self._transform_cards(activity)
        return activity

    # ============================================================ internals

    async def _acquire_teams_token(self, ctx: TurnContext) -> str:
        """Use core's Authorization to get a Teams-scoped token.

        Authorization stays in core; Teams Extension just consumes it via
        the host pointer captured at initialize().
        """
        # Pseudocode — wire to host.auth or core's Connections lookup:
        # provider = self._host.auth.get_token_provider(ctx.identity, ctx.activity.service_url)
        # return await provider.get_access_token(audience, scopes)
        raise NotImplementedError

    def _parse_channel_data(self, activity: Activity) -> dict:
        return dict(activity.channel_data or {})

    def _inject_channel_data(self, ctx: TurnContext, activity: Activity) -> Activity:
        # Merge anything devs set on ctx.turn_state["teams.outbound_channel_data"]
        extra = ctx.turn_state.get("teams.outbound_channel_data") or {}
        if extra:
            base = dict(activity.channel_data or {})
            base.update(extra)
            activity = activity.model_copy(update={"channel_data": base})
        return activity

    def _inject_mentions(self, activity: Activity) -> Activity:
        # Walk activity.entities for "mention" types added via TeamsContext.mention(...)
        return activity

    def _transform_cards(self, activity: Activity) -> Activity:
        # Adaptive Card → Teams-flavored attachment variants live here.
        return activity

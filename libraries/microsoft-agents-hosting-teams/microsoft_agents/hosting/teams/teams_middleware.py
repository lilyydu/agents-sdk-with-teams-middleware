"""Teams plumbing as middleware. Zero core changes required.

This is the *only* thing the developer wires into the adapter to get Teams-
correct behavior. It runs on every turn (reactive, proactive, continue_
conversation) because all of those flows go through the same middleware
pipeline (`ChannelServiceAdapter.run_pipeline`).

Two responsibilities, both implemented using public core APIs that exist
today:

1. Inbound enrichment ("on_context_created"):
     Right after core builds the TurnContext, we attach Teams-specific
     helpers (TeamsConnectorClient, TeamsInfo, parsed channelData).

2. Outbound transforms ("on_activity_sending"):
     Registered via the existing `TurnContext.on_send_activities` callback
     list. Runs before each outbound activity reaches the connector.
     Rewrites mentions, injects notification metadata, transforms cards.

Filter: everything here is gated on ``activity.channel_id == "msteams"``.
On any other channel the middleware is a no-op.
"""

from __future__ import annotations

from typing import Awaitable, Callable

from microsoft_agents.activity import Activity, ResourceResponse
from microsoft_agents.hosting.core.middleware_set import Middleware
from microsoft_agents.hosting.core.turn_context import TurnContext

from .connector.teams.teams_connector_client import TeamsConnectorClient
from .teams_info import TeamsInfo

TEAMS_CHANNEL_ID = "msteams"

# turn_state keys — public so handlers can look these up if they want
TEAMS_INFO_KEY = "teams.info"
TEAMS_DATA_KEY = "teams.channel_data"
CONNECTOR_CLIENT_KEY = "ConnectorClient"  # core's convention


class TeamsMiddleware(Middleware):
    """Owns all Teams-specific behavior. Lives in the Teams package."""

    async def on_turn(
        self,
        context: TurnContext,
        logic: Callable[[], Awaitable],
    ) -> None:
        if context.activity.channel_id != TEAMS_CHANNEL_ID:
            await logic()
            return

        # ── on_context_created equivalent ──
        self._swap_connector_client(context)
        self._attach_teams_helpers(context)
        self._parse_channel_data(context)

        # ── on_activity_sending equivalent ──
        # Register a per-turn outbound interceptor. Core invokes this for
        # every Activity that leaves through ctx.send_activity / send_activities.
        context.on_send_activities(self._on_send_activities)

        await logic()

    # ───────────────────────── inbound ─────────────────────────

    def _swap_connector_client(self, ctx: TurnContext) -> None:
        """Ensure outbound traffic uses TeamsConnectorClient, not the generic one."""
        # TODO: build with the right credentials / scopes for Teams.
        # Today the factory picks this based on recipient.role; this hook
        # makes the choice explicit and channel-owned.
        existing = ctx.turn_state.get(CONNECTOR_CLIENT_KEY)
        if isinstance(existing, TeamsConnectorClient):
            return
        # If we need credentials, pull them off the existing client.
        # Concrete construction is filled in during implementation.
        # ctx.turn_state[CONNECTOR_CLIENT_KEY] = TeamsConnectorClient(...)

    def _attach_teams_helpers(self, ctx: TurnContext) -> None:
        """Expose Teams API helpers on the turn so handlers don't construct them."""
        ctx.turn_state.setdefault(TEAMS_INFO_KEY, TeamsInfo)  # class, not instance — most methods are class methods

    def _parse_channel_data(self, ctx: TurnContext) -> None:
        """Make activity.channel_data easy to read."""
        raw = ctx.activity.channel_data or {}
        ctx.turn_state[TEAMS_DATA_KEY] = raw  # TODO: replace with a typed model

    # ───────────────────────── outbound ────────────────────────

    async def _on_send_activities(
        self,
        ctx: TurnContext,
        activities: list[Activity],
        next: Callable[[], Awaitable[list[ResourceResponse]]],
    ) -> list[ResourceResponse]:
        for activity in activities:
            self._inject_notification(activity)
            self._format_mentions(activity)
            self._transform_card(activity)
        return await next()

    @staticmethod
    def _inject_notification(activity: Activity) -> None:
        """Make Teams ping the user instead of silently delivering."""
        if activity.channel_data is None:
            activity.channel_data = {}
        activity.channel_data.setdefault("notification", {"alert": True})

    @staticmethod
    def _format_mentions(activity: Activity) -> None:
        """Translate '@Name' references into Teams <at>Name</at> + entity."""
        # TODO: scan activity.text for known mentions, rewrite to <at> tags,
        # and append Mention entities to activity.entities.
        return

    @staticmethod
    def _transform_card(activity: Activity) -> None:
        """Add Teams-specific properties to Adaptive Cards (e.g. msteams: { width })."""
        # TODO: walk activity.attachments for adaptive cards; merge Teams quirks.
        return

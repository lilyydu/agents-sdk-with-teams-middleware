"""Extension protocols for the Agents-for-python host.

Two distinct seams, intentionally separated:

* `ChannelAdapter`  — *plumbing*. Runs for every turn on its channel,
                       regardless of whether the dev wrote any channel-specific
                       handlers. Use to swap the connector client, enrich the
                       TurnContext (Teams API client, Graph client, parsed
                       channelData) and transform outbound activities.

* `DeepExtension`   — *surface*. Opt-in, additive. Registers rich
                       channel-specific handlers (invokes, dialog, card actions,
                       reactions). Builds a rich context for its handlers
                       without disturbing the generic cross-channel handlers.

A package may ship either, both, or neither. `Extension` is a convenience
bundle that lets a package expose one entry-point to `app.use(...)`.

Versioning: every extension declares `name`, `version`, and `requires_core`
(SemVer range). The host enforces compatibility at `app.use(...)` time and
fails fast on mismatch.
"""

from __future__ import annotations

from typing import Optional, Protocol, runtime_checkable

from microsoft_agents.activity import Activity

# These are real symbols in microsoft-agents-hosting-core. Imports are written
# as if this file lives inside that package.
from .turn_context import TurnContext


@runtime_checkable
class ChannelAdapter(Protocol):
    """Channel-scoped plumbing. Two hooks: in and out."""

    name: str
    version: str
    requires_core: str  # SemVer range, e.g. ">=2.0,<3.0"

    def owns_channel(self, channel_id: str) -> bool:
        """True if this adapter applies to activities from `channel_id`."""
        ...

    async def on_context_created(self, ctx: TurnContext) -> None:
        """Called by core right after a TurnContext is built.

        Both reactive (HTTP /api/messages) and proactive
        (`continue_conversation`) paths invoke this. Typical work:

          * swap `ctx.connector_client` for a channel-specific one
          * attach channel-specific helper clients (Teams API, Graph)
          * parse `activity.channel_data` into a typed structure on ctx
        """
        ...

    async def on_activity_sending(
        self, ctx: TurnContext, activity: Activity
    ) -> Activity:
        """Called by core before each outbound activity leaves the process.

        Return the (possibly rewritten) activity. Typical work:

          * inject `channelData` (notifications, importance, etc.)
          * add `<at>` mention entities
          * transform Adaptive Cards to Teams-specific variants
          * chunk for Teams streaming size limits
        """
        ...


@runtime_checkable
class DeepExtension(Protocol):
    """Channel-scoped rich handler surface. Opt-in."""

    name: str
    version: str
    requires_core: str

    def owns_channel(self, channel_id: str) -> bool: ...

    def can_handle(self, ctx: TurnContext) -> bool:
        """True if this extension wants to dispatch this turn itself.

        First MATCH wins per turn. Returning False lets the generic
        cross-platform handlers run (LCM path).
        """
        ...

    async def handle(self, ctx: TurnContext) -> None:
        """Dispatch the turn to a registered handler.

        Implementations typically:
          1. Build a rich channel context (e.g. `TeamsContext`) from `ctx`.
          2. Match the activity against the extension's internal route table.
          3. Invoke the developer's handler.
          4. For invoke turns, set the invoke response on `ctx`.
        """
        ...


class Extension(Protocol):
    """A bundle a package may export for a single `app.use(...)` call.

    Either field may be None. Typical Teams package exposes both.
    """

    channel_adapter: Optional[ChannelAdapter]
    deep_extension: Optional[DeepExtension]

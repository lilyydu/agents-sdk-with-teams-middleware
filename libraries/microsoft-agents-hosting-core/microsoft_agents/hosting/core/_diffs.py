"""Annotated diffs needed in existing core files (DOCS, not runtime code).

Shows the exact call sites in existing files where the new ExtensionRegistry
plugs in. Each block is the smallest change that wires the new contract into
the existing pipelines.

Files to edit in the real PR:

  1. libraries/microsoft-agents-hosting-core/microsoft_agents/hosting/core/
       app/agent_application.py
  2. libraries/microsoft-agents-hosting-core/microsoft_agents/hosting/core/
       channel_service_adapter.py
  3. libraries/microsoft-agents-hosting-core/microsoft_agents/hosting/core/
       app/proactive/proactive.py

No file is moved. No existing call is removed. Three additive call sites.
"""

# ---------------------------------------------------------------------------
# 1) AgentApplication.use() — registration entry point
# ---------------------------------------------------------------------------
# File: app/agent_application.py
#
# Add to imports:
#     from ..extension import ChannelAdapter, DeepExtension, Extension
#     from ..extension_registry import ExtensionRegistry
#
# Add to __init__ (around the other field initializations near line 81):
#     self.extensions = ExtensionRegistry()
#
# Add new public method (next to .use_authorization etc.):
#
#     def use(self, ext) -> "AgentApplication":
#         """Register a ChannelAdapter, DeepExtension, or Extension bundle."""
#         _check_version(self.core_version, ext.requires_core)  # SemVer guard
#         if isinstance(ext, ChannelAdapter):
#             self.extensions.register_channel_adapter(ext)
#         if isinstance(ext, DeepExtension):
#             self.extensions.register_deep_extension(ext)
#         if isinstance(ext, Extension):
#             if ext.channel_adapter:
#                 self.extensions.register_channel_adapter(ext.channel_adapter)
#             if ext.deep_extension:
#                 self.extensions.register_deep_extension(ext.deep_extension)
#         if hasattr(ext, "initialize"):
#             ext.initialize(self)
#         return self


# ---------------------------------------------------------------------------
# 2) ChannelServiceAdapter.process_activity() — invoke on_context_created
# ---------------------------------------------------------------------------
# File: channel_service_adapter.py  (line ~368)
#
# Existing flow constructs context: TurnContext(...) and stashes
# turn_state["ConnectorClient"]. Right BEFORE running the agent's on_turn,
# call the matching ChannelAdapter's hook so it can swap the client and enrich.
#
# ADD (after context construction, before middleware/agent dispatch):
#
#     channel_adapter = self._app.extensions.channel_adapter_for(activity.channel_id)
#     if channel_adapter is not None:
#         await channel_adapter.on_context_created(context)


# ---------------------------------------------------------------------------
# 3) ChannelServiceAdapter.send_activities() — invoke on_activity_sending
# ---------------------------------------------------------------------------
# File: channel_service_adapter.py  (line ~57)
#
# Wherever outbound activities are dispatched to connector_client.conversations,
# wrap each item with the channel adapter's transform hook.
#
# CHANGE (in the loop that sends each activity):
#
#     channel_adapter = self._app.extensions.channel_adapter_for(
#         context.activity.channel_id
#     )
#     for activity in activities:
#         if channel_adapter is not None:
#             activity = await channel_adapter.on_activity_sending(context, activity)
#         # ... existing reply_to_activity / send_to_conversation call ...
#
# All outbound traffic — including from DeepExtension handlers and TeamsContext
# — funnels through here, so the transformation hook always runs.


# ---------------------------------------------------------------------------
# 4) Proactive.continue_conversation() — same on_context_created hook
# ---------------------------------------------------------------------------
# File: app/proactive/proactive.py  (line ~225)
#
# After building the proactive TurnContext but before running middleware/callback:
#
#     channel_adapter = self._app.extensions.channel_adapter_for(ref.channel_id)
#     if channel_adapter is not None:
#         await channel_adapter.on_context_created(context)
#
# Same hook, both entry points → no drift between reactive and proactive.


# ---------------------------------------------------------------------------
# 5) Route dispatch — DeepExtension wins for its channel
# ---------------------------------------------------------------------------
# File: app/agent_application.py — wherever routes are matched.
#
# BEFORE iterating the generic route table:
#
#     deep = self.extensions.deep_extension_for(context.activity.channel_id)
#     if deep is not None and deep.can_handle(context):
#         await deep.handle(context)
#         return
#     # else: fall through to generic route list (LCM path)
#
# Note: the existing AgentApplication.add_route(...) system continues to work
# for both generic handlers and for routes the DeepExtension itself registers
# internally. The new step here is the explicit first-match check against
# `deep_extension_for(channel_id)` so the Teams team controls Teams traffic.

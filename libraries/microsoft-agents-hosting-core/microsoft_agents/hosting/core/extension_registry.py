"""Registry: resolves ChannelAdapter / DeepExtension by channel_id.

Used by:
  * CloudAdapter.process()             → channel_adapter_for(activity.channel_id)
  * Proactive.continue_conversation()  → channel_adapter_for(ref.channel_id)
  * AgentApplication route dispatch    → deep_extension_for(activity.channel_id)

Lookup is first-match-wins in registration order; tiebreak is deterministic.
"""

from __future__ import annotations

from typing import Optional

from .extension import ChannelAdapter, DeepExtension


class ExtensionRegistry:
    def __init__(self) -> None:
        self._channel_adapters: list[ChannelAdapter] = []
        self._deep_extensions: list[DeepExtension] = []

    # ----- registration -----

    def register_channel_adapter(self, adapter: ChannelAdapter) -> None:
        self._channel_adapters.append(adapter)

    def register_deep_extension(self, ext: DeepExtension) -> None:
        self._deep_extensions.append(ext)

    # ----- lookup -----

    def channel_adapter_for(
        self, channel_id: Optional[str]
    ) -> Optional[ChannelAdapter]:
        if not channel_id:
            return None
        for a in self._channel_adapters:
            if a.owns_channel(channel_id):
                return a
        return None

    def deep_extension_for(
        self, channel_id: Optional[str]
    ) -> Optional[DeepExtension]:
        if not channel_id:
            return None
        for d in self._deep_extensions:
            if d.owns_channel(channel_id):
                return d
        return None

    # ----- introspection (for diagnostics + tests) -----

    @property
    def channel_adapters(self) -> list[ChannelAdapter]:
        return list(self._channel_adapters)

    @property
    def deep_extensions(self) -> list[DeepExtension]:
        return list(self._deep_extensions)

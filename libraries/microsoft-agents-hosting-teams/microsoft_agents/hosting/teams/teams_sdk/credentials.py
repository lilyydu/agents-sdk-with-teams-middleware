"""

Adapt Agents SDK's ``MsalConnectionManager`` to teams.py's token-callback shape.

teams.py's ``App`` constructor accepts a ``token`` callable of shape
``(scope, tenant_id) -> str | Awaitable[str]`` which it wraps into
``TokenCredentials`` internally. Agents SDK already has a working MSAL
provider; this module produces the callback so developers don't have to
duplicate token logic.

Tenant awareness:
    When invoked inside a turn handled by ``TeamsSDKMiddleware``, the
    callback reads the inbound ``ClaimsIdentity`` from
    ``turn_state[AGENT_IDENTITY_KEY]`` and asks the connection manager
    for the tenant-scoped provider via
    ``get_token_provider(claims_identity, service_url)``.

    Outside a turn (proactive callbacks, background tasks, startup hooks)
    the ContextVar is unset and we fall back to ``get_default_connection()``.
"""

from __future__ import annotations

import inspect
import logging
from typing import Any, Awaitable, Callable, Optional, Union

from microsoft_agents.hosting.core.channel_adapter import ChannelAdapter

from ._context import _agent_turn_context

logger = logging.getLogger(__name__)


def make_agent_sdk_token_provider(
    connection_manager: Any,
) -> Callable[[Union[str, list[str]], Optional[str]], Awaitable[str]]:
    """Return a token callback that delegates to a ConnectionManager.

    Args:
        connection_manager: An Agents SDK ConnectionManager (typically an
            instance of MsalConnectionManager). Must expose
            ``get_default_connection()`` and
            ``get_token_provider(claims_identity, service_url)`` returning
            an AccessTokenProviderBase with
            ``get_access_token(resource_url, scopes, force_refresh=False)``.

    Returns:
        An async callable matching teams.py's ``TokenCredentials.token``
        signature. Tenant-aware when invoked inside a turn handled by
        ``TeamsSDKMiddleware``; falls back to the default connection for
        proactive / background calls.
    """

    async def _token(
        scope: Union[str, list[str]],
        tenant_id: Optional[str] = None,  # noqa: ARG001 — tenant flows via ClaimsIdentity
    ) -> str:
        scopes = [scope] if isinstance(scope, str) else list(scope)
        # Strip ".default" off the first scope to derive the resource_url
        # MSAL wants — '.default' is appended back internally.
        first = scopes[0] if scopes else ""
        resource_url = first[: -len("/.default")] if first.endswith("/.default") else first

        provider = _select_provider(connection_manager, resource_url)
        result = provider.get_access_token(resource_url, scopes)
        if inspect.isawaitable(result):
            result = await result
        return result

    return _token


def _select_provider(connection_manager: Any, service_url: str) -> Any:
    """
    Pick a tenant-scoped provider when a turn is in scope, else default.

    Looks up the inbound ClaimsIdentity from the current turn's
    ``turn_state[AGENT_IDENTITY_KEY]``. If anything is missing — no turn
    in scope, no identity stashed, or the connection manager rejects the
    lookup — fall back to the default connection. That fallback is the
    only path used by proactive sends and non-turn callers.
    """
    context = _agent_turn_context.get(None)
    if context is None:
        return connection_manager.get_default_connection()

    claims_identity = context.turn_state.get(ChannelAdapter.AGENT_IDENTITY_KEY)
    if claims_identity is None:
        return connection_manager.get_default_connection()

    target_url = service_url or (context.activity.service_url or "")
    if not target_url:
        return connection_manager.get_default_connection()

    try:
        return connection_manager.get_token_provider(claims_identity, target_url)
    except Exception as exc:  # noqa: BLE001 — degrade gracefully on lookup failure
        logger.debug(
            "tenant-scoped token provider lookup failed (%s); using default connection",
            exc,
        )
        return connection_manager.get_default_connection()

"""

Adapt Agents SDK's ``MsalConnectionManager`` to teams.py's token-callback shape.

teams.py's ``App`` constructor accepts a ``token`` callable of shape
``(scope, tenant_id) -> str | Awaitable[str]`` which it wraps into
``TokenCredentials`` internally. Agents SDK already has a working MSAL
provider; this module produces the callback so developers don't have to
duplicate token logic.

Connection-aware:
    When invoked inside a turn handled by ``TeamsSDKMiddleware``, the
    callback reads the inbound ``ClaimsIdentity`` from
    ``turn_state[AGENT_IDENTITY_KEY]`` and asks the connection manager
    for the matching connection via
    ``get_token_provider(claims_identity, service_url)``. This matters
    when more than one connection is registered (e.g. different app
    registrations per channel or skill audience); with a single
    connection it returns the same provider as
    ``get_default_connection()``.

    Note: this does NOT swap tenants per user. Outbound Bot Framework
    Service tokens are always acquired against the bot's home tenant
    configured on the connection itself; the inbound identity only
    selects WHICH connection to use, not the tenant inside that
    connection.

    Outside a turn (proactive callbacks, background tasks, startup hooks)
    the ContextVar is unset and we fall back to ``get_default_connection()``.

Agentic (Agent 365) requests:
    When teams.py needs an outbound token to act as the *agentic user*
    (Teams-channel agentic turns, i.e. ``recipient.role == "agenticUser"``),
    it invokes this callback with an ``agentic_identity``. In that case we
    mint the token via the connection's ``get_agentic_user_token`` — the same
    Agents SDK path as ``AgenticUserAuthorization`` — instead of the normal
    bot token. The bot's own client_id is the blueprint the token derives
    from; teams.py supplies the agentic app-instance id, user id, and scopes.
    Agentic *bot* (app-instance) outbound is intentionally not handled here:
    teams.py has no such path, so use ``context.send_activity`` via
    ``agent_sdk_turn_context`` for that case.
"""

from __future__ import annotations

import inspect
import logging
from typing import Any, Awaitable, Callable, Optional, Union

from microsoft_agents.hosting.core.channel_adapter import ChannelAdapter

from ._context import _agent_sdk_turn_context

logger = logging.getLogger(__name__)

# Routing key used to select the connection that mints agentic tokens,
# mirroring the Agents SDK's ``get_token_provider(identity, "agentic")``
# convention (and the ``CONNECTIONSMAP ... SERVICEURL=agentic`` entry). When no
# dedicated agentic connection is registered the lookup falls back to the
# default connection, which is sufficient for the single-connection case.
_AGENTIC_ROUTING_KEY = "agentic"


def make_agent_sdk_token_provider(
    connection_manager: Any,
) -> Callable[..., Awaitable[str]]:
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
        signature. Connection-aware (picks among multiple registered
        connections when applicable) when invoked inside a turn handled
        by ``TeamsSDKMiddleware``; falls back to the default connection
        for proactive / background calls. When teams.py passes an
        ``agentic_identity`` (agentic-user turns) it mints an agentic user
        token via ``get_agentic_user_token`` instead of the bot token.
    """

    async def _token(
        scope: Union[str, list[str]],
        tenant_id: Optional[str] = None,
        *,
        agentic_identity: Optional[Any] = None,
    ) -> str:
        scopes = [scope] if isinstance(scope, str) else list(scope)

        # Agentic-user path: teams.py passes an ``agentic_identity`` when the
        # outbound call must act as the agentic user rather than the bot. Mint
        # the user token through the same Agents SDK connection (single source
        # of truth), mirroring ``AgenticUserAuthorization``. teams.py only
        # populates ``agentic_identity`` for agentic turns; it is ``None``
        # otherwise, in which case we fall through to the normal bot token.
        agentic_user_id = getattr(agentic_identity, "agentic_user_id", None)
        if agentic_identity is not None and agentic_user_id:
            provider = _select_provider(connection_manager, _AGENTIC_ROUTING_KEY)
            token = await provider.get_agentic_user_token(
                getattr(agentic_identity, "tenant_id", None) or tenant_id,
                agentic_identity.agentic_app_id,
                agentic_user_id,
                scopes,
            )
            if not token:
                raise RuntimeError(
                    "Agents SDK connection returned no agentic user token; the "
                    "connection must be an MSAL provider with agentic support."
                )
            return token

        # Non-agentic path: outbound Bot Framework Service token. Strip
        # ".default" off the first scope to derive the resource_url MSAL wants
        # — '.default' is appended back internally.
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
    Pick the matching connection when a turn is in scope, else default.

    Looks up the inbound ClaimsIdentity from the current turn's
    ``turn_state[AGENT_IDENTITY_KEY]``. If anything is missing — no turn
    in scope, no identity stashed, or the connection manager rejects the
    lookup — fall back to the default connection. That fallback is the
    only path used by proactive sends and non-turn callers.

    NOTE: This does NOT perform per-user-tenant token acquisition. The
    chosen connection's own tenant_id is always used for the outbound
    token. This call only matters when more than one connection is
    registered.
    """
    context = _agent_sdk_turn_context.get(None)
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
            "connection lookup failed (%s); using default connection",
            exc,
        )
        return connection_manager.get_default_connection()

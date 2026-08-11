"""One-call factory: build a teams.py App from an Agents SDK app + connections.

This is the recommended entry point. It owns the three things that otherwise
have to be wired by hand at every call site:

    1. Credential extraction — pulls ``client_id``/``tenant_id`` from the
       connection manager's default configuration.
    2. Token-provider wiring — bridges teams.py's outbound token callback to
       the Agents SDK ``MsalConnectionManager``.
    3. Middleware install — registers ``TeamsSDKMiddleware`` on the Agents
       SDK adapter so Teams turns are short-circuited to teams.py.

Returns a fully constructed ``microsoft_teams.apps.App`` ready for handler
registration.
"""

from __future__ import annotations

from typing import Any, Callable, Optional

from microsoft_agents.hosting.core.app.agent_application import AgentApplication
from microsoft_agents.hosting.core.authorization.connections import Connections
from microsoft_agents.hosting.core.turn_context import TurnContext

from microsoft_teams.apps import App

from .credentials import make_agent_sdk_token_provider
from .middleware import TeamsSDKMiddleware


def use_teams_sdk(
    app: AgentApplication,
    connection_manager: Connections,
    should_bypass_teams: Optional[Callable[[TurnContext], bool]] = None,
    **teams_app_kwargs: Any,
) -> App:
    """Wire teams.py into ``app`` and return the configured ``App``.

    Args:
        app: The Agents SDK application whose adapter will own the HTTP
            endpoint. ``TeamsSDKMiddleware`` is installed on ``app.adapter``.
        connection_manager: Source of credentials and tokens. The default
            connection's ``CLIENT_ID`` and ``TENANT_ID`` are used to
            construct the teams.py ``App``; its token providers are wrapped
            so teams.py's outbound calls use the same credentials.
        should_bypass_teams: Optional predicate evaluated only for
            Teams-channel turns. Return ``True`` to bypass teams.py routing and
            force the turn to fall through to the Agents SDK even when teams.py
            has a matching route.
        **teams_app_kwargs: Extra keyword arguments forwarded to
            ``microsoft_teams.apps.App``. Use this for ``logger``, ``plugins``,
            or any other ``App`` constructor option. ``client_id``,
            ``tenant_id``, and ``token`` are reserved and will raise if
            passed here.

    Returns:
        The configured teams.py ``App``. Register handlers on the returned
        object (``@teams_app.on_message_pattern(...)``, etc.).

    After this call:
        * Teams turns claimed by ``should_bypass_teams`` → fall through to ``app``.
        * Teams turns with a matching teams.py handler → handled by the App.
        * Teams turns with no match → fall through to ``app``'s handlers.
        * Any other channel → handled by ``app`` unchanged.
    """
    reserved = {"client_id", "tenant_id", "token"} & teams_app_kwargs.keys()
    if reserved:
        raise TypeError(
            f"use_teams_sdk owns {sorted(reserved)}; remove from teams_app_kwargs."
        )

    auth = connection_manager.get_default_connection_configuration()
    teams_app = App(
        client_id=auth.CLIENT_ID,
        tenant_id=auth.TENANT_ID,
        token=make_agent_sdk_token_provider(connection_manager),
        **teams_app_kwargs,
    )

    app.adapter.use(TeamsSDKMiddleware(teams_app, should_bypass_teams))
    return teams_app

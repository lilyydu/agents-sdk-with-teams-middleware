from __future__ import annotations

from contextvars import ContextVar

from microsoft_agents.hosting.core.turn_context import TurnContext

_agent_sdk_turn_context: ContextVar[TurnContext] = ContextVar(
    "microsoft_agents.hosting.teams.teams_sdk.agent_sdk_turn_context"
)


def agent_sdk_turn_context() -> TurnContext:
    """Return the Agents SDK ``TurnContext`` for the current turn.

    :raises LookupError: If called outside the scope of a Teams SDK-handled
        turn (e.g. from a startup hook, a background task that doesn't
        inherit the turn's context, or a non-Teams turn that fell through).
    """
    return _agent_sdk_turn_context.get()


__all__ = ["agent_sdk_turn_context"]

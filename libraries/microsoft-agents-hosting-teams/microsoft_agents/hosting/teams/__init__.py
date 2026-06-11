from .teams_activity_handler import TeamsActivityHandler
from .teams_agent_extension import (
    TeamsAgentExtension,
    MessageExtension,
    TaskModule,
    Meeting,
)
from .teams_info import TeamsInfo

# RFC: Teams SDK integration (embeds the teams.py App via middleware)
from .teams_sdk import (
    TeamsSDKMiddleware,
    agent_sdk_turn_context,
    make_agent_sdk_token_provider,
    use_teams_sdk,
)

__all__ = [
    "TeamsActivityHandler",
    "TeamsAgentExtension",
    "MessageExtension",
    "TaskModule",
    "Meeting",
    "TeamsInfo",
    # Teams SDK integration
    "TeamsSDKMiddleware",
    "agent_sdk_turn_context",
    "make_agent_sdk_token_provider",
    "use_teams_sdk",
]

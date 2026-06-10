from __future__ import annotations

from microsoft_teams.cards.core import (
    AdaptiveCard,
    ExecuteAction,
    Fact,
    FactSet,
    SubmitAction,
    TextBlock,
    TextInput,
)
from microsoft_teams.cards.utilities import OpenDialogData


def help_card() -> AdaptiveCard:
    """List the commands the sample understands, grouped by which app owns them."""
    return AdaptiveCard(
        version="1.5",
        body=[
            TextBlock(text="Teams SDK Feature Showcase", weight="Bolder", size="Large"),
            TextBlock(
                text="Teams SDK handlers (TEAMS_APP)",
                weight="Bolder",
                spacing="Medium",
            ),
            FactSet(
                facts=[
                    Fact(title="help", value="This command list"),
                    Fact(title="cards", value="Adaptive Card with Action.Execute invoke"),
                    Fact(title="citation", value="AI labels, citations, sensitivity, feedback"),
                    Fact(title="stream", value="Streaming response with informative updates"),
                    Fact(title="react", value="Bot adds/removes emoji reactions"),
                    Fact(title="quote", value="Bot quotes its own message"),
                    Fact(title="proactive", value="Delayed proactive message"),
                    Fact(title="task", value="Task module fetch/submit flow"),
                    Fact(title="turn context", value="Use Agent SDK TurnContext from a teams.py handler"),
                ]
            ),
            TextBlock(
                text="Agents SDK fallthrough handlers (AGENT_APP)",
                weight="Bolder",
                spacing="Medium",
            ),
            FactSet(
                facts=[
                    Fact(title="agents react", value="Reach teams.py's API client from an Agents SDK handler"),
                    Fact(title="agents proactive", value="Trigger a proactive send from an Agents SDK handler"),
                    Fact(title="agents citation", value="Build a Teams citation activity from an Agents SDK handler"),
                    Fact(title="anything else", value="Echo via Agents SDK '[Agent SDK] You said: ...'"),
                ]
            ),
        ],
    )


def ping_card() -> AdaptiveCard:
    """Demo card whose button fires an Action.Execute invoke (verb=ping)."""
    return AdaptiveCard(
        version="1.5",
        body=[
            TextBlock(text="🎯 Invoke demo", weight="Bolder", size="Medium"),
            TextBlock(
                text="Press the button to fire an Action.Execute invoke.",
                wrap=True,
            ),
        ],
        actions=[
            ExecuteAction(
                # teams.py's @on_card_action_execute('ping') matches on
                # data.action (NOT on verb). Both are sent for parity.
                verb="ping",
                title="Ping the bot",
                data={"action": "ping", "sentAt": "now"},
            ),
        ],
    )


def task_launcher_card() -> AdaptiveCard:
    """Card whose button opens a task module via the task/fetch invoke."""
    return AdaptiveCard(
        version="1.5",
        body=[
            TextBlock(text="📋 Task module demo", weight="Bolder"),
            TextBlock(
                text="Press the button to open a task module.",
                wrap=True,
            ),
        ],
        actions=[
            SubmitAction(title="Open task module").with_data(
                OpenDialogData("open_task")
            ),
        ],
    )


def task_form_card() -> AdaptiveCard:
    """The form shown inside the task module."""
    return AdaptiveCard(
        version="1.5",
        body=[
            TextBlock(text="Tell us something:", weight="Bolder"),
            TextInput(id="note").with_placeholder("Type here…"),
        ],
        actions=[SubmitAction(title="Submit")],
    )

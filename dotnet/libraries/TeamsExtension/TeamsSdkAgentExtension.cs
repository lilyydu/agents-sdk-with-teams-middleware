// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Agents.Builder;
using Microsoft.Agents.Builder.App;
using Microsoft.Agents.Builder.State;
using Microsoft.Agents.Core.Models;
using Microsoft.Agents.Core.Serialization;
using Microsoft.Extensions.Logging;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Schema;
using Microsoft.Teams.Core.Schema;
using System;
using System.Threading;
using System.Threading.Tasks;
using TeamsInvokeResponse = Microsoft.Teams.Apps.Handlers.InvokeResponse;

namespace TeamsExtension;

/// <summary>
/// AgentApplication extension that routes matching Teams turns to a <see cref="TeamsBotApplication"/>
/// using <see cref="AgentApplication.OnBeforeTurn(Microsoft.Agents.Builder.TurnEventHandler)"/>.
/// </summary>
public sealed class TeamsSdkAgentExtension : AgentExtension
{
    private static readonly AsyncLocal<ITurnContext?> _currentTurnContext = new();
    private readonly TeamsBotApplication _teamsBot;
    private readonly ILogger<TeamsSdkAgentExtension>? _logger;

    public TeamsSdkAgentExtension(
        AgentApplication agentApplication,
        TeamsBotApplication teamsBot,
        ILogger<TeamsSdkAgentExtension>? logger = null)
    {
        _teamsBot = teamsBot ?? throw new ArgumentNullException(nameof(teamsBot));
        _logger = logger;
        ChannelId = Channels.Msteams;

        agentApplication.OnBeforeTurn(RouteTeamsTurnAsync);
    }

    /// <summary>
    /// The Agent SDK <see cref="ITurnContext"/> for the current Teams SDK-routed turn.
    /// </summary>
    public static ITurnContext? CurrentTurnContext => _currentTurnContext.Value;

    /// <summary>
    /// Returns the current Agent SDK <see cref="ITurnContext"/> or throws when unavailable.
    /// </summary>
    public static ITurnContext RequireTurnContext()
        => _currentTurnContext.Value
           ?? throw new InvalidOperationException(
               "RequireTurnContext() called outside a TeamsSdkAgentExtension-routed turn.");

    private async Task<bool> RouteTeamsTurnAsync(
        ITurnContext turnContext,
        ITurnState turnState,
        CancellationToken cancellationToken)
    {
        if (turnContext.Activity.ChannelId != Channels.Msteams)
        {
            return true;
        }

        string activityJson = ProtocolJsonSerializer.ToJson(turnContext.Activity);

        // HasMatchingRoute mutates CoreActivity, so use a dedicated route-check instance.
        CoreActivity routeCheckActivity = CoreActivity.FromJsonString(activityJson);
        if (!_teamsBot.HasMatchingRoute(routeCheckActivity))
        {
            _logger?.LogDebug(
                "TeamsSdkAgentExtension: no matching Teams SDK route for activity {ActivityId}, falling through to AgentApplication",
                turnContext.Activity.Id);
            return true;
        }

        _logger?.LogDebug(
            "TeamsSdkAgentExtension: routing msteams activity {ActivityId} to Teams SDK",
            turnContext.Activity.Id);

        _currentTurnContext.Value = turnContext;
        try
        {
            CoreActivity coreActivity = CoreActivity.FromJsonString(activityJson);

            if (turnContext.Activity.Type == ActivityTypes.Invoke)
            {
                TeamsInvokeResponse invokeResponse = await _teamsBot.ProcessInvokeAsync(coreActivity, cancellationToken).ConfigureAwait(false);
                if (invokeResponse is not null)
                {
                    var responseActivity = Activity.CreateInvokeResponseActivity(invokeResponse.Body, invokeResponse.Status);
                    await turnContext.SendActivityAsync((Activity)responseActivity, cancellationToken).ConfigureAwait(false);
                }
            }
            else if (_teamsBot.OnActivity is not null)
            {
                await _teamsBot.OnActivity(coreActivity, cancellationToken).ConfigureAwait(false);
            }

            // Do not continue AgentApplication route evaluation when Teams SDK handled it.
            return false;
        }
        finally
        {
            _currentTurnContext.Value = null;
        }
    }
}

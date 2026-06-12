// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Agents.Builder;
using Microsoft.Agents.Core.Models;
using Microsoft.Agents.Core.Serialization;
using Microsoft.Extensions.Logging;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Schema;
using Microsoft.Teams.Core.Schema;
using System;
using System.Threading;
using System.Threading.Tasks;
using IMiddleware = Microsoft.Agents.Builder.IMiddleware;
using TeamsInvokeResponse = Microsoft.Teams.Apps.Handlers.InvokeResponse;

namespace TeamsSdk;

/// <summary>
/// Middleware that intercepts incoming activities and routes Teams-channel
/// traffic to a <see cref="Microsoft.Teams.Apps.TeamsBotApplication"/> instead of letting it
/// continue through the Agent SDK pipeline.
/// </summary>
/// <remarks>
/// Register this as a singleton <see cref="IMiddleware"/> in DI.  Because the
/// <see cref="Microsoft.Agents.Hosting.AspNetCore.CloudAdapter"/> constructor
/// takes <c>IMiddleware[]</c> (and .NET DI does not auto-resolve array types),
/// you must also register <c>IMiddleware[]</c> explicitly — see <c>Program.cs</c>.
/// </remarks>
public class TeamsSdkMiddleware : IMiddleware
{
    /// <summary>
    /// The Agent SDK <see cref="ITurnContext"/> for the current turn.
    /// Uses <see cref="AsyncLocal{T}"/> so it flows within the same async context
    /// regardless of which thread the turn executes on (non-invoke activities are
    /// processed on a background thread where HttpContext is unavailable).
    /// </summary>
    public static ITurnContext? CurrentTurnContext => _currentTurnContext.Value;

    /// <summary>
    /// Returns the Agent SDK <see cref="ITurnContext"/> for the current turn, throwing if
    /// none is in scope.  Use this from Teams SDK handlers when the turn context is required;
    /// use <see cref="CurrentTurnContext"/> when its absence is a valid, handled state.
    /// </summary>
    /// <returns>The non-null <see cref="ITurnContext"/> for the current turn.</returns>
    /// <exception cref="InvalidOperationException">
    /// Thrown when called outside a Teams SDK middleware turn.  This usually means the caller
    /// is running on a detached async task that did not inherit the execution context (a
    /// <see cref="System.Threading.Timer"/> callback,
    /// <see cref="System.Threading.ThreadPool"/>.UnsafeQueueUserWorkItem, a raw
    /// <see cref="System.Threading.Thread"/>, or code where <c>ExecutionContext</c> flow was
    /// suppressed) or outside the request lifecycle entirely.
    /// </exception>
    public static ITurnContext RequireTurnContext()
        => _currentTurnContext.Value
           ?? throw new InvalidOperationException(
               "RequireTurnContext() called outside a Teams SDK middleware turn. " +
               "This usually means the caller is running on a detached async task that did " +
               "not inherit the execution context (a Timer callback, " +
               "ThreadPool.UnsafeQueueUserWorkItem, a raw Thread, or code where " +
               "ExecutionContext flow was suppressed) or outside the request lifecycle entirely.");

    private static readonly AsyncLocal<ITurnContext?> _currentTurnContext = new();

    private readonly TeamsBotApplication _teamsBot;
    private readonly ILogger<TeamsSdkMiddleware> _logger;

    public TeamsSdkMiddleware(TeamsBotApplication teamsBot, ILogger<TeamsSdkMiddleware> logger)
    {
        _teamsBot = teamsBot ?? throw new ArgumentNullException(nameof(teamsBot));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    public async Task OnTurnAsync(ITurnContext turnContext, NextDelegate next, CancellationToken cancellationToken = default)
    {
        if (turnContext.Activity.ChannelId == Channels.Msteams)
        {
            // Bridge: serialize the Agent SDK IActivity to JSON, then deserialize
            // into the Teams SDK activity model.  Both SDKs implement the same
            // Activity Protocol wire format, so the conversion is lossless.
            string activityJson = ProtocolJsonSerializer.ToJson(turnContext.Activity);

            // HasMatchingRoute calls TeamsActivity.FromActivity which mutates the
            // CoreActivity (Extract removes entries from Properties). Deserialize a
            // separate copy for the match check so the real activity stays intact.
            CoreActivity routeCheckActivity = CoreActivity.FromJsonString(activityJson);

            // Only route to Teams SDK if a registered handler matches this activity.
            // Unmatched activities fall through to the Agent SDK pipeline.
            if (_teamsBot.HasMatchingRoute(routeCheckActivity))
            {
                _logger.LogDebug("TeamsSdkMiddleware: routing msteams activity {ActivityId} to Teams SDK", turnContext.Activity.Id);

                // Deserialize a fresh CoreActivity for the handler (the routeCheckActivity
                // was mutated by HasMatchingRoute's Extract calls).
                CoreActivity coreActivity = CoreActivity.FromJsonString(activityJson);

                // Make the Agent SDK turn context available to Teams SDK handlers.
                // Non-invoke activities are processed on a background thread where
                // HttpContext is unavailable, so we use an AsyncLocal that flows
                // within the same async context regardless of thread.  Handlers
                // access it via TeamsSdkMiddleware.CurrentTurnContext.
                _currentTurnContext.Value = turnContext;

                if (turnContext.Activity.Type == ActivityTypes.Invoke)
                {
                    // Invoke activities require special handling: Teams SDK returns an
                    // InvokeResponse that must be bridged into Agent SDK's StackState
                    // so that ProcessTurnResults can write the correct HTTP response.
                    // Using ProcessInvokeAsync avoids the double-write conflict that
                    // occurs when OnActivity writes directly to HttpContext.Response.
                    TeamsInvokeResponse invokeResponse = await _teamsBot.ProcessInvokeAsync(coreActivity, cancellationToken);

                    if (invokeResponse is not null)
                    {
                        var responseActivity = Activity.CreateInvokeResponseActivity(invokeResponse.Body, invokeResponse.Status);
                        await turnContext.SendActivityAsync((Activity)responseActivity, cancellationToken);
                    }
                }
                else
                {
                    // Non-invoke activities: use the standard OnActivity delegate which
                    // sends responses via ConversationClient.
                    if (_teamsBot.OnActivity != null)
                    {
                        await _teamsBot.OnActivity(coreActivity, cancellationToken);
                    }
                }

                // Short-circuit: do NOT call next() — Teams SDK handled this activity.
                return;
            }

            _logger.LogDebug("TeamsSdkMiddleware: no matching Teams SDK route for activity {ActivityId}, falling through to Agent SDK", turnContext.Activity.Id);
        }

        // Non-Teams channels (or unmatched Teams activities) continue to the Agent SDK pipeline.
        await next(cancellationToken);
    }
}

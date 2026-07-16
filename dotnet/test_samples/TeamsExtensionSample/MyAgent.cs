// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Agents.Builder;
using Microsoft.Agents.Builder.App;
using Microsoft.Agents.Builder.State;
using Microsoft.Agents.Core.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Teams.Apps.Handlers;
using Microsoft.Teams.Apps.Schema;
using Microsoft.Teams.Apps.Schema.Entities;
using System;
using System.Threading;
using System.Threading.Tasks;
using TeamsExtension;

namespace TeamsMiddlewareSample;

/// <summary>
/// Agent SDK handler that demonstrates using Teams SDK services via
/// <see cref="MyTeamsBot"/> (which extends <c>TeamsBotApplication</c>)
/// from within Agent SDK turn handlers.  Messages that don't match any
/// Teams SDK route in <see cref="MyTeamsBot"/> fall through extension routing
/// to this agent.
/// </summary>
[Agent(name: "MyAgent", description: "Agent SDK handler with Teams SDK integration", version: "1.0")]
public class MyAgent : AgentApplication
{
    private readonly MyTeamsBot _teamsBot;
    private readonly ILogger<MyAgent> _logger;

    public MyAgent(AgentApplicationOptions options, MyTeamsBot teamsBot, ILogger<MyAgent> logger)
        : base(options)
    {
        _teamsBot = teamsBot;
        _logger = logger;
        this.UseTeamsExtension(teamsBot);

        OnConversationUpdate(ConversationUpdateEvents.MembersAdded, WelcomeMessageAsync);
        OnActivity(ActivityTypes.Message, OnMessageAsync, rank: RouteRank.Last);
    }

    private async Task WelcomeMessageAsync(ITurnContext turnContext, ITurnState turnState, CancellationToken cancellationToken)
    {
        foreach (ChannelAccount member in turnContext.Activity.MembersAdded)
        {
            if (member.Id != turnContext.Activity.Recipient.Id)
            {
                await turnContext.SendActivityAsync(MessageFactory.Text("Hello from Agent SDK!"), cancellationToken);
            }
        }
    }

    private async Task OnMessageAsync(ITurnContext turnContext, ITurnState turnState, CancellationToken cancellationToken)
    {
        string? text = turnContext.Activity.RemoveRecipientMention()?.Trim();

        if (string.Equals(text, "agents react", StringComparison.OrdinalIgnoreCase))
        {
            await HandleAgentsReactAsync(turnContext, cancellationToken);
            return;
        }

        if (string.Equals(text, "agents proactive", StringComparison.OrdinalIgnoreCase))
        {
            await HandleAgentsProactiveAsync(turnContext, cancellationToken);
            return;
        }

        if (string.Equals(text, "agents citation", StringComparison.OrdinalIgnoreCase))
        {
            await HandleAgentsCitationAsync(turnContext, cancellationToken);
            return;
        }

        // Default echo
        await turnContext.SendActivityAsync($"[Agent SDK] You said: {turnContext.Activity.Text}", cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Sends a message via Agent SDK, then uses <see cref="MyTeamsBot.Api"/>
    /// (scoped to the service URL) to add and remove a reaction.
    /// </summary>
    private async Task HandleAgentsReactAsync(ITurnContext turnContext, CancellationToken cancellationToken)
    {
        var response = await turnContext.SendActivityAsync(
            MessageFactory.Text("[Agent SDK] Sent this message — now using Teams SDK to react..."),
            cancellationToken);

        if (response?.Id != null)
        {
            string conversationId = turnContext.Activity.Conversation.Id;
            var api = _teamsBot.Api.ForServiceUrl(new Uri(turnContext.Activity.ServiceUrl));

            await Task.Delay(1500, cancellationToken);
            await api.Conversations.Reactions.AddAsync(
                conversationId, response.Id, ReactionTypes.Like, cancellationToken: cancellationToken);

            await Task.Delay(2000, cancellationToken);
            await api.Conversations.Reactions.DeleteAsync(
                conversationId, response.Id, ReactionTypes.Like, cancellationToken: cancellationToken);

            await turnContext.SendActivityAsync(
                MessageFactory.Text("[Agent SDK] Reaction added and removed via TeamsBotApplication.Api!"),
                cancellationToken);
        }
    }

    /// <summary>
    /// Uses <see cref="MyTeamsBot.SendAsync"/> to send a proactive message
    /// outside the current turn, directly through the Teams SDK.
    /// </summary>
    private async Task HandleAgentsProactiveAsync(ITurnContext turnContext, CancellationToken cancellationToken)
    {
        string conversationId = turnContext.Activity.Conversation.Id;

        await turnContext.SendActivityAsync(
            MessageFactory.Text("[Agent SDK] You will receive a proactive message via TeamsBotApplication.SendAsync in ~3 seconds..."),
            cancellationToken);

        _ = Task.Run(async () =>
        {
            await Task.Delay(3000);
            try
            {
                await _teamsBot.SendAsync(conversationId,
                    "[Teams SDK] Proactive message sent from Agent SDK handler via TeamsBotApplication.SendAsync!");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to send proactive message via Teams SDK");
            }
        });
    }

    /// <summary>
    /// Builds a citation message using Teams SDK types, then sends it
    /// through <see cref="MyTeamsBot.Api"/> from an Agent SDK handler.
    /// </summary>
    private async Task HandleAgentsCitationAsync(ITurnContext turnContext, CancellationToken cancellationToken)
    {
        // Build a rich Teams message using Teams SDK types
        var teamsMessage = TeamsActivity.CreateBuilder()
            .WithText("Agent SDK handler built this citation using Teams SDK types [1].")
            .AddCitation(1, new CitationAppearance
            {
                Name = "Agent SDK Documentation",
                Abstract = "The Agent SDK provides the hosting infrastructure for this sample.",
                Url = new Uri("https://learn.microsoft.com/en-us/microsoft-365/agents-sdk/"),
                Icon = CitationIcon.Text
            })
            .AddAIGenerated()
            .AddFeedback()
            .Build();

        // Send via TeamsBotApplication.Api scoped to the current service URL
        string conversationId = turnContext.Activity.Conversation.Id;
        var api = _teamsBot.Api.ForServiceUrl(new Uri(turnContext.Activity.ServiceUrl));
        await api.Conversations.Activities.CreateAsync(conversationId, teamsMessage, cancellationToken: cancellationToken);

        _logger.LogInformation("Citation message sent via TeamsBotApplication.Api from Agent SDK handler");
    }
}

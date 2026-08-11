// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

﻿#pragma warning disable ExperimentalTeamsQuotedReplies // ReplyAsync quotes the inbound message.
#pragma warning disable ExperimentalTeamsTargeted       // WithRecipient(targeted) is experimental.

using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Api.Clients;
using Microsoft.Teams.Apps.Handlers;
using Microsoft.Teams.Apps.Handlers.TaskModules;
using Microsoft.Teams.Apps.Schema;
using Microsoft.Teams.Core.Schema;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;

namespace TeamsMiddlewareSample;

/// <summary>
/// Teams SDK routes that own matching Teams turns before the Agents SDK sees them.
/// </summary>
public class MyTeamsBot : TeamsBotApplication
{
    public MyTeamsBot(ApiClient api, IHttpContextAccessor accessor, ILogger<MyTeamsBot> logger, TeamsBotApplicationOptions? options = null)
        : base(api, accessor, logger, options)
    {
        this.OnMessage("help", async (context, ct) =>
        {
            var attachment = TeamsAttachment.CreateBuilder()
                .WithAdaptiveCard(ParseCard("""
                    {
                      "type": "AdaptiveCard",
                      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                      "version": "1.5",
                      "body": [
                        {
                          "type": "TextBlock",
                          "text": "Teams SDK Feature Showcase",
                          "weight": "Bolder",
                          "size": "Large",
                          "wrap": true
                        },
                        {
                          "type": "TextBlock",
                          "text": "Teams SDK handlers (MyTeamsBot)",
                          "weight": "Bolder",
                          "spacing": "Medium"
                        },
                        {
                          "type": "FactSet",
                          "facts": [
                            { "title": "help", "value": "This command list" },
                            { "title": "react", "value": "Bot adds/removes emoji reactions" },
                            { "title": "quote", "value": "Bot quotes your message" },
                            { "title": "targeted", "value": "Ephemeral message visible only to sender" },
                            { "title": "task", "value": "Task module fetch/submit flow" }
                          ]
                        },
                        {
                          "type": "TextBlock",
                          "text": "Agents SDK fallthrough handlers (MyAgent)",
                          "weight": "Bolder",
                          "spacing": "Medium"
                        },
                        {
                          "type": "FactSet",
                          "facts": [
                            { "title": "help", "value": "Plain-text help on non-Teams channels" },
                            { "title": "channel", "value": "Report the channel and how it was routed" },
                            { "title": "whoami", "value": "Graph profile via OAuth connection graphuser" },
                            { "title": "mail", "value": "Recent mail via OAuth connection graphmail" },
                            { "title": "signout", "value": "Clear both OAuth handler caches" },
                            { "title": "agents sdk react", "value": "Reach Teams reactions API from the Agent SDK" },
                            { "title": "agents sdk proactive", "value": "Send via Teams SDK API client from the Agent SDK" },
                            { "title": "anything else", "value": "Echo via the Agent SDK" }
                          ]
                        }
                      ]
                    }
                    """))
                .Build();

            await context.SendActivityAsync(new MessageActivity([attachment]), ct);
        });

        this.OnMessage("react", async (context, ct) =>
        {
            var response = await context.SendAsync("React to this message! I'll add thumbs-up and remove it.", ct);
            if (response?.Id is null)
            {
                return;
            }

            string conversationId = context.Activity.Conversation!.Id;
            await Task.Delay(2000, ct);
            await context.Api.Conversations.Reactions.AddAsync(conversationId, response.Id, ReactionTypes.Like, cancellationToken: ct);
            await Task.Delay(2000, ct);
            await context.Api.Conversations.Reactions.DeleteAsync(conversationId, response.Id, ReactionTypes.Like, cancellationToken: ct);
        });

        this.OnMessage("quote", async (context, ct) =>
        {
            await context.ReplyAsync("Quoting your message!", ct);
        });

        this.OnMessage("targeted", async (context, ct) =>
        {
            var sender = context.Activity.From;
            var targeted = new MessageActivity("👁️ This message is only visible to you.")
                .WithRecipient(new ConversationAccount { Id = sender!.Id, Name = sender.Name }, isTargeted: true);
            await context.SendActivityAsync(targeted, ct);
        });

        this.OnMessage("task", async (context, ct) =>
        {
            var attachment = TeamsAttachment.CreateBuilder()
                .WithAdaptiveCard(ParseCard("""
                    {
                      "type": "AdaptiveCard",
                      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                      "version": "1.5",
                      "body": [
                        {
                          "type": "TextBlock",
                          "text": "Task module demo",
                          "weight": "Bolder",
                          "size": "Medium"
                        },
                        {
                          "type": "TextBlock",
                          "text": "Press the button to open a task module.",
                          "wrap": true
                        }
                      ],
                      "actions": [
                        {
                          "type": "Action.Submit",
                          "title": "Open task module",
                          "data": {
                            "msteams": { "type": "task/fetch" }
                          }
                        }
                      ]
                    }
                    """))
                .Build();

            await context.SendActivityAsync(new MessageActivity([attachment]), ct);
        });

        this.OnTaskFetch(async (context, ct) =>
        {
            var attachment = TeamsAttachment.CreateBuilder()
                .WithAdaptiveCard(ParseCard("""
                    {
                      "type": "AdaptiveCard",
                      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                      "version": "1.5",
                      "body": [
                        {
                          "type": "TextBlock",
                          "text": "Task Module Form",
                          "weight": "Bolder",
                          "size": "Medium"
                        },
                        {
                          "type": "Input.Text",
                          "id": "note",
                          "placeholder": "Type here...",
                          "label": "Your response"
                        }
                      ],
                      "actions": [
                        {
                          "type": "Action.Submit",
                          "title": "Submit"
                        }
                      ]
                    }
                    """))
                .Build();

            return TaskModuleResponse.CreateBuilder()
                .WithType(TaskModuleResponseType.Continue)
                .WithTitle("Sample Task Module")
                .WithCard(attachment)
                .WithHeight(TaskModuleSize.Medium)
                .WithWidth(TaskModuleSize.Medium)
                .Build();
        });

        this.OnTaskSubmit(async (context, ct) =>
        {
            await context.SendAsync($"[Teams SDK] Task module submitted. Data: {context.Activity.Value?.Data}", ct);
            return TaskModuleResponse.CreateBuilder()
                .WithType(TaskModuleResponseType.Message)
                .WithMessage("Done.")
                .Build();
        });

        this.OnMessageReaction(async (context, ct) =>
        {
            var added = context.Activity.ReactionsAdded?.Select(reaction => reaction.Type ?? string.Empty).ToArray() ?? [];
            var removed = context.Activity.ReactionsRemoved?.Select(reaction => reaction.Type ?? string.Empty).ToArray() ?? [];
            await context.SendAsync($"[Teams SDK] Reactions: added=[{string.Join(", ", added)}] removed=[{string.Join(", ", removed)}]", ct);
        });
    }

    private static JsonElement ParseCard(string json)
    {
        using JsonDocument document = JsonDocument.Parse(json);
        return document.RootElement.Clone();
    }
}

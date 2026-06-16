// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

#pragma warning disable ExperimentalTeamsQuotedReplies // Quote is experimental
#pragma warning disable ExperimentalTeamsTargeted       // WithRecipient(targeted) is experimental

using Microsoft.Agents.Builder;
using Microsoft.Agents.Builder.State;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Api.Clients;
using Microsoft.Teams.Apps.Handlers;
using Microsoft.Teams.Apps.Handlers.MessageExtension;
using Microsoft.Teams.Apps.Handlers.TaskModules;
using Microsoft.Teams.Apps.Schema;
using Microsoft.Teams.Apps.Schema.Entities;
using Microsoft.Teams.Core.Schema;
using TeamsSdk;
using System;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace TeamsMiddlewareSample;

/// <summary>
/// Comprehensive Teams SDK feature showcase.
/// Each command demonstrates a different Teams SDK capability.
/// </summary>
public class MyTeamsBot : TeamsBotApplication
{
    private readonly ILogger<MyTeamsBot> _logger;
    private readonly ConversationState _conversationState;

    public MyTeamsBot(ApiClient api, IHttpContextAccessor accessor, ILogger<MyTeamsBot> logger, ConversationState conversationState, TeamsBotApplicationOptions? options = null)
        : base(api, accessor, logger, options)
    {
        _logger = logger;
        _conversationState = conversationState;

        // ── help ──────────────────────────────────────────────────────
        this.OnMessage("help", async (context, ct) =>
        {
            var card = JsonDocument.Parse("""
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
                            "type": "FactSet",
                            "facts": [
                                { "title": "help", "value": "This command list" },
                                { "title": "cards", "value": "Adaptive Card with Action.Execute invoke" },
                                { "title": "citation", "value": "AI labels, citations, sensitivity, feedback" },
                                { "title": "stream", "value": "Streaming response with informative updates" },
                                { "title": "react", "value": "Bot adds/removes emoji reactions" },
                                { "title": "quote", "value": "Bot quotes its own message" },
                                { "title": "targeted", "value": "Ephemeral message visible only to sender" },
                                { "title": "proactive", "value": "Delayed proactive message" },
                                { "title": "task", "value": "Task module fetch/submit flow" },
                                { "title": "turn context", "value": "Use Agent SDK ITurnContext from Teams SDK handler" },
                                { "title": "(search box)", "value": "Message extension query — search from compose box" },
                                { "title": "(action cmd)", "value": "Message extension action — form via task module" }
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "text": "Agent SDK handlers (using Teams SDK services)",
                            "weight": "Bolder",
                            "size": "Medium",
                            "spacing": "Large",
                            "wrap": true
                        },
                        {
                            "type": "FactSet",
                            "facts": [
                                { "title": "agents react", "value": "Add/remove reactions via Teams SDK ApiClient" },
                                { "title": "agents proactive", "value": "Proactive message via TeamsBotApplication.SendAsync" },
                                { "title": "agents citation", "value": "Citation built with Teams SDK types, sent via ConversationClient" }
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "text": "Also reacts to: message reactions, feedback buttons, meeting start/end",
                            "wrap": true,
                            "size": "Small",
                            "isSubtle": true,
                            "spacing": "Medium"
                        }
                    ]
                }
                """).RootElement;

            var attachment = TeamsAttachment.CreateBuilder()
                .WithAdaptiveCard(card)
                .Build();

            await context.SendActivityAsync(new MessageActivity([attachment]), ct);
        });

        // ── cards ─────────────────────────────────────────────────────
        this.OnMessage("cards", async (context, ct) =>
        {
            var card = JsonDocument.Parse("""
                {
                    "type": "AdaptiveCard",
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "version": "1.5",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": "Invoke Test Card",
                            "weight": "Bolder",
                            "size": "Medium"
                        },
                        {
                            "type": "Input.Text",
                            "id": "userInput",
                            "placeholder": "Type something here",
                            "label": "Your input"
                        }
                    ],
                    "actions": [
                        {
                            "type": "Action.Execute",
                            "title": "Submit",
                            "data": { "action": "test" }
                        }
                    ]
                }
                """).RootElement;

            var attachment = TeamsAttachment.CreateBuilder()
                .WithAdaptiveCard(card)
                .Build();

            await context.SendActivityAsync(new MessageActivity([attachment]), ct);
        });

        // ── citation ──────────────────────────────────────────────────
        this.OnMessage("citation", async (context, ct) =>
        {
            var activity = TeamsActivity.CreateBuilder()
                .WithText("Here is a response with citations [1].")
                .AddCitation(1, new CitationAppearance
                {
                    Name = "Teams SDK Documentation",
                    Abstract = "The Teams SDK provides building blocks for agents in Microsoft Teams.",
                    Url = new Uri("https://learn.microsoft.com/en-us/microsoftteams/"),
                    Icon = CitationIcon.Text,
                    UsageInfo = new SensitiveUsageEntity { Name = "Confidential" }
                })
                .AddAIGenerated()
                .AddFeedback()
                .Build();

            await context.SendActivityAsync(activity, ct);
        });

        // ── stream ────────────────────────────────────────────────────
        this.OnMessage("stream", async (context, ct) =>
        {
            var writer = TeamsStreamingWriter.CreateFromContext(context);

            await writer.SendInformativeUpdateAsync("Thinking...", ct);
            await Task.Delay(1000, ct);

            await writer.AppendResponseAsync("Streaming is a powerful feature ", ct);
            await Task.Delay(800, ct);

            await writer.AppendResponseAsync("that lets you send incremental updates ", ct);
            await Task.Delay(800, ct);

            await writer.AppendResponseAsync("to users in real time.", ct);
            await Task.Delay(500, ct);

            var finalMessage = new MessageActivity("Streaming is a powerful feature that lets you send incremental updates to users in real time.");
            finalMessage.AddAIGenerated();
            finalMessage.AddFeedback();

            await writer.FinalizeResponseAsync(finalMessage, ct);
        });

        // ── react ─────────────────────────────────────────────────────
        this.OnMessage("react", async (context, ct) =>
        {
            var response = await context.SendAsync("React to this message! (bot will add a thumbs-up then remove it)", ct);
            if (response?.Id != null)
            {
                string conversationId = context.Activity.Conversation!.Id;

                await Task.Delay(1500, ct);
                await context.Api.Conversations.Reactions.AddAsync(conversationId, response.Id, ReactionTypes.Like, cancellationToken: ct);

                await Task.Delay(2000, ct);
                await context.Api.Conversations.Reactions.DeleteAsync(conversationId, response.Id, ReactionTypes.Like, cancellationToken: ct);

                await context.SendAsync("Thumbs-up added then removed!", ct);
            }
        });

        // ── quote ─────────────────────────────────────────────────────
        this.OnMessage("quote", async (context, ct) =>
        {
            var response = await context.SendAsync("This is the original message.", ct);
            if (response?.Id != null)
            {
                await Task.Delay(1000, ct);
                await context.Quote(response.Id, "Quoting my own message!", ct);
            }
        });

        // ── targeted ─────────────────────────────────────────────────
        this.OnMessage("targeted", async (context, ct) =>
        {
            var sender = context.Activity.From;
            var targeted = new MessageActivity("👁️ This message is only visible to you.")
                .WithRecipient(new ConversationAccount { Id = sender!.Id, Name = sender.Name }, isTargeted: true);
            await context.SendActivityAsync(targeted, ct);
        });

        // ── proactive ─────────────────────────────────────────────────
        this.OnMessage("proactive", async (context, ct) =>
        {
            string conversationId = context.Activity.Conversation!.Id;
            await context.SendAsync("You will receive a proactive message in ~3 seconds...", ct);

            _ = Task.Run(async () =>
            {
                await Task.Delay(3000);
                try
                {
                    await SendAsync(conversationId, "This is a proactive message sent outside the turn!");
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to send proactive message");
                }
            });
        });

        // ── task ──────────────────────────────────────────────────────
        this.OnMessage("task", async (context, ct) =>
        {
            var card = JsonDocument.Parse("""
                {
                    "type": "AdaptiveCard",
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "version": "1.5",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": "Click the button to open a task module",
                            "weight": "Bolder",
                            "size": "Medium"
                        }
                    ],
                    "actions": [
                        {
                            "type": "Action.Submit",
                            "title": "Open Task Module",
                            "data": {
                                "msteams": { "type": "task/fetch" }
                            }
                        }
                    ]
                }
                """).RootElement;

            var attachment = TeamsAttachment.CreateBuilder()
                .WithAdaptiveCard(card)
                .Build();

            await context.SendActivityAsync(new MessageActivity([attachment]), ct);
        });

        // ── turn context ──────────────────────────────────────────────
        // Use Agents SDK ConversationState from inside a Teams handler.
        // Demonstrates the key value prop: Teams handlers can tap into Agents SDK
        // infrastructure (state management, storage) without duplicating it.
        this.OnMessage("turn context", async (context, ct) =>
        {
            var agentCtx = TeamsSdkMiddleware.CurrentTurnContext;
            if (agentCtx is null)
            {
                await context.SendAsync("[Teams SDK] Agent SDK turn context is not available.", ct);
                return;
            }

            // Access Agents SDK conversation state from inside a Teams handler
            await _conversationState.LoadAsync(agentCtx, cancellationToken: ct);
            int count = _conversationState.GetValue<int>("messageCount", () => 0);
            count++;
            _conversationState.SetValue("messageCount", count);
            await _conversationState.SaveChangesAsync(agentCtx, cancellationToken: ct);

            await context.SendAsync(
                $"📊 This conversation has received **{count}** message(s) " +
                "(tracked via Agents SDK ConversationState from a Teams handler).", ct);
        });

        // ── Adaptive Card Action handler ──────────────────────────────
        this.OnAdaptiveCardAction(async (context, ct) =>
        {
            var action = context.Activity.Value?.Action;
            string info = action?.Data != null
                ? JsonSerializer.Serialize(action.Data)
                : "(no data)";

            await context.SendAsync($"[Teams SDK] Adaptive Card action received. Data: {info}", ct);
            return InvokeResponse.Ok();
        });

        // ── Task Module: fetch ────────────────────────────────────────
        this.OnTaskFetch(async (context, ct) =>
        {
            var formCard = JsonDocument.Parse("""
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
                            "id": "taskInput",
                            "placeholder": "Enter something",
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
                """).RootElement;

            var attachment = TeamsAttachment.CreateBuilder()
                .WithAdaptiveCard(formCard)
                .Build();

            return TaskModuleResponse.CreateBuilder()
                .WithType(TaskModuleResponseType.Continue)
                .WithTitle("Sample Task Module")
                .WithCard(attachment)
                .WithHeight(TaskModuleSize.Medium)
                .WithWidth(TaskModuleSize.Medium)
                .Build();
        });

        // ── Task Module: submit ───────────────────────────────────────
        this.OnTaskSubmit(async (context, ct) =>
        {
            var data = context.Activity.Value?.Data;
            await context.SendAsync($"[Teams SDK] Task module submitted. Data: {data}", ct);

            return TaskModuleResponse.CreateBuilder()
                .WithType(TaskModuleResponseType.Message)
                .WithMessage("Task module completed successfully!")
                .Build();
        });

        // ── Message Extension: query ───────────────────────────────────
        this.OnQuery(async (context, ct) =>
        {
            var query = context.Activity.Value;
            string searchText = query?.Parameters?.FirstOrDefault()?.Value ?? "";

            // initialRun sends "true" as the parameter value — treat as empty
            if (string.Equals(searchText, "true", StringComparison.OrdinalIgnoreCase))
                searchText = "";

            _logger.LogInformation("Message extension query: CommandId={CommandId}, Search={Search}",
                query?.CommandId, searchText);

            // Return sample results as hero cards
            string[] items = ["Alpha", "Beta", "Gamma", "Delta", "Epsilon"];
            var results = items
                .Where(i => string.IsNullOrEmpty(searchText) || i.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                .Select(item => TeamsAttachment.CreateBuilder()
                    .WithContentType(AttachmentContentType.HeroCard)
                    .WithContent(new
                    {
                        title = item,
                        subtitle = string.IsNullOrEmpty(searchText) ? "All results" : $"Match for '{searchText}'",
                        text = $"This is a demo search result: {item}"
                    })
                    .Build())
                .ToArray();

            if (results.Length == 0)
            {
                return MessageExtensionResponse.CreateBuilder()
                    .WithType(MessageExtensionResponseType.Message)
                    .WithText($"No results found for '{searchText}'")
                    .Build();
            }

            return MessageExtensionResponse.CreateBuilder()
                .WithType(MessageExtensionResponseType.Result)
                .WithAttachmentLayout(TeamsAttachmentLayout.List)
                .WithAttachments(results)
                .Build();
        });

        // ── Message Extension: action (fetch task) ────────────────────
        this.OnFetchTask(async (context, ct) =>
        {
            var action = context.Activity.Value;
            _logger.LogInformation("Message extension fetch task: CommandId={CommandId}", action?.CommandId);

            var formCard = JsonDocument.Parse("""
                {
                    "type": "AdaptiveCard",
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "version": "1.5",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": "Message Extension Action",
                            "weight": "Bolder",
                            "size": "Medium"
                        },
                        {
                            "type": "Input.Text",
                            "id": "title",
                            "placeholder": "Enter a title",
                            "label": "Title"
                        },
                        {
                            "type": "Input.Text",
                            "id": "description",
                            "placeholder": "Enter a description",
                            "label": "Description",
                            "isMultiline": true
                        }
                    ],
                    "actions": [
                        {
                            "type": "Action.Submit",
                            "title": "Submit"
                        }
                    ]
                }
                """).RootElement;

            var attachment = TeamsAttachment.CreateBuilder()
                .WithAdaptiveCard(formCard)
                .Build();

            return MessageExtensionActionResponse.CreateBuilder()
                .WithTask(
                    TaskModuleResponse.CreateBuilder()
                        .WithType(TaskModuleResponseType.Continue)
                        .WithTitle("Create Item")
                        .WithCard(attachment)
                        .WithHeight(TaskModuleSize.Medium)
                        .WithWidth(TaskModuleSize.Medium))
                .Build();
        });

        // ── Message Extension: action (submit) ───────────────────────
        this.OnSubmitAction(async (context, ct) =>
        {
            var action = context.Activity.Value;
            _logger.LogInformation("Message extension submit action: CommandId={CommandId}, Data={Data}",
                action?.CommandId, action?.Data);

            // Return a result card that gets inserted into the compose box
            var resultCard = TeamsAttachment.CreateBuilder()
                .WithContentType(AttachmentContentType.HeroCard)
                .WithContent(new
                {
                    title = "Item Created",
                    subtitle = "Via message extension action",
                    text = $"Data: {action?.Data}"
                })
                .Build();

            return MessageExtensionActionResponse.CreateBuilder()
                .WithComposeExtension(
                    MessageExtensionResponse.CreateBuilder()
                        .WithType(MessageExtensionResponseType.Result)
                        .WithAttachmentLayout(TeamsAttachmentLayout.List)
                        .WithAttachments(resultCard))
                .Build();
        });

        // ── Message Reactions ─────────────────────────────────────────
        this.OnMessageReaction(async (context, ct) =>
        {
            var added = context.Activity.ReactionsAdded;
            var removed = context.Activity.ReactionsRemoved;

            string summary = "";
            if (added?.Count > 0)
                summary += $"Reactions added: {string.Join(", ", added.Select(r => r.Type))}. ";
            if (removed?.Count > 0)
                summary += $"Reactions removed: {string.Join(", ", removed.Select(r => r.Type))}. ";

            if (!string.IsNullOrEmpty(summary))
                await context.SendAsync($"[Teams SDK] {summary.Trim()}", ct);
        });

        // ── Feedback ──────────────────────────────────────────────────
        this.OnMessageSubmitFeedback(async (context, ct) =>
        {
            var feedback = context.Activity.Value;
            _logger.LogInformation("Feedback received: {Value}", feedback);
            await context.SendAsync("[Teams SDK] Thanks for your feedback!", ct);
            return InvokeResponse.Ok();
        });

        // ── Members Added (welcome) ──────────────────────────────────
        this.OnMembersAdded(async (context, ct) =>
        {
            await context.SendAsync("Hello from Teams SDK! Type **help** to see available commands.", ct);
        });

        // ── Meeting Start ─────────────────────────────────────────────
        this.OnMeetingStart(async (context, ct) =>
        {
            _logger.LogInformation("Meeting started");
            await context.SendAsync("[Teams SDK] Meeting has started!", ct);
        });

        // ── Meeting End ───────────────────────────────────────────────
        this.OnMeetingEnd(async (context, ct) =>
        {
            _logger.LogInformation("Meeting ended");
            await context.SendAsync("[Teams SDK] Meeting has ended!", ct);
        });

    }
}

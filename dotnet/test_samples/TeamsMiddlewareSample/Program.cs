// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Agents.Builder.State;
using Microsoft.Agents.Hosting.AspNetCore;
using Microsoft.Agents.Storage;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using TeamsMiddlewareSample;
using TeamsSdk;

WebApplicationBuilder builder = WebApplication.CreateBuilder(args);

// ── Agent SDK ──────────────────────────────────────────────────────
builder.AddAgent<MyAgent>();
builder.Services.AddSingleton<IStorage, MemoryStorage>();
builder.Services.AddSingleton<ConversationState>();
builder.Services.AddAgentAspNetAuthentication(builder.Configuration);

// ── Teams SDK ──────────────────────────────────────────────────────
// One call: registers MyTeamsBot + its Teams API/auth chain (via AgentSdkAuthHandler)
// and installs the routing middleware on the CloudAdapter pipeline.
builder.Services.AddTeamsSdk<MyTeamsBot>();

WebApplication app = builder.Build();

app.UseAuthentication();
app.UseAuthorization();

app.MapAgentRootEndpoint();
app.MapAgentApplicationEndpoints(requireAuth: !app.Environment.IsDevelopment());

app.Run();

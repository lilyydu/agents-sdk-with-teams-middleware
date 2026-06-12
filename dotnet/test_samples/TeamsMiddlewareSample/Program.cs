// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System.Linq;
using Microsoft.Agents.Builder;
using Microsoft.Agents.Hosting.AspNetCore;
using Microsoft.Agents.Storage;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using TeamsMiddlewareSample;
using TeamsSdk;
using IMiddleware = Microsoft.Agents.Builder.IMiddleware;

WebApplicationBuilder builder = WebApplication.CreateBuilder(args);

// ── Agent SDK ──────────────────────────────────────────────────────
builder.AddAgent<MyAgent>();
builder.Services.AddSingleton<IStorage, MemoryStorage>();
builder.Services.AddAgentAspNetAuthentication(builder.Configuration);

// ── Teams SDK (reuses Agent SDK auth via AgentSdkAuthHandler) ────
builder.Services.AddTeamsSdkWithAgentAuth<MyTeamsBot>();

// ── Routing Middleware ─────────────────────────────────────────────
// CloudAdapter's constructor takes an optional IMiddleware[] parameter.
// .NET DI does not auto-resolve array types, so we register the array
// explicitly after all individual IMiddleware registrations.
builder.Services.AddSingleton<IMiddleware, TeamsSdkMiddleware>();
builder.Services.AddSingleton<IMiddleware[]>(sp =>
    sp.GetServices<IMiddleware>().ToArray());

WebApplication app = builder.Build();

app.UseAuthentication();
app.UseAuthorization();

app.MapAgentRootEndpoint();
app.MapAgentApplicationEndpoints(requireAuth: !app.Environment.IsDevelopment());

app.Run();

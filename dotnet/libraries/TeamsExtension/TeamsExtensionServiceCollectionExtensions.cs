// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Api.Clients;
using Microsoft.Teams.Core;
using System;
using System.Net.Http;

namespace TeamsExtension;

/// <summary>
/// Extension methods for registering Teams SDK services using Agent SDK auth.
/// </summary>
public static class TeamsExtensionServiceCollectionExtensions
{
    private const string HttpClientName = "TeamsBot";

    /// <summary>
    /// Registers Teams SDK services and a <see cref="TeamsBotApplication"/> subclass for use with
    /// the AgentApplication extension bridge.
    /// </summary>
    /// <typeparam name="T">A <see cref="TeamsBotApplication"/> subclass.</typeparam>
    public static IServiceCollection AddTeamsExtension<T>(this IServiceCollection services)
        where T : TeamsBotApplication
    {
        services.AddHttpContextAccessor();

        services.AddTransient<AgentSdkAuthHandler>();
        services.AddHttpClient(HttpClientName)
            .AddHttpMessageHandler<AgentSdkAuthHandler>();

        services.AddSingleton<ConversationClient>(sp =>
        {
            var httpClient = sp.GetRequiredService<IHttpClientFactory>().CreateClient(HttpClientName);
            return new ConversationClient(httpClient, sp.GetRequiredService<ILogger<ConversationClient>>());
        });

        services.AddSingleton<UserTokenClient>(sp =>
        {
            var httpClient = sp.GetRequiredService<IHttpClientFactory>().CreateClient(HttpClientName);
            return new UserTokenClient(
                httpClient,
                sp.GetRequiredService<IConfiguration>(),
                sp.GetRequiredService<ILogger<UserTokenClient>>());
        });

        services.AddSingleton<ApiClient>(sp =>
        {
            var httpClient = sp.GetRequiredService<IHttpClientFactory>().CreateClient(HttpClientName);
            return new ApiClient(
                httpClient,
                sp.GetRequiredService<ConversationClient>(),
                sp.GetRequiredService<UserTokenClient>());
        });

        services.AddSingleton<T>(sp => (T)ActivatorUtilities.CreateInstance(sp, typeof(T)));
        services.AddSingleton<TeamsBotApplication>(sp => sp.GetRequiredService<T>());

        return services;
    }
}

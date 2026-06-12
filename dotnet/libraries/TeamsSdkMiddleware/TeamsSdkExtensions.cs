// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System.Net.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Api.Clients;
using Microsoft.Teams.Core;

namespace TeamsSdk;

/// <summary>
/// Extension methods for registering the Teams SDK services using the Agent SDK's
/// authentication (<c>IConnections</c>) instead of a separate AzureAd config section.
/// </summary>
public static class TeamsSdkExtensions
{
    private const string HttpClientName = "TeamsBot";

    /// <summary>
    /// Registers a <see cref="TeamsBotApplication"/> subclass and its dependencies
    /// (<see cref="ApiClient"/>, <see cref="ConversationClient"/>,
    /// <see cref="UserTokenClient"/>) using a named <see cref="HttpClient"/> whose
    /// outbound requests are authenticated by <see cref="AgentSdkAuthHandler"/>.
    /// </summary>
    /// <typeparam name="T">A <see cref="TeamsBotApplication"/> subclass.</typeparam>
    public static IServiceCollection AddTeamsSdkWithAgentAuth<T>(this IServiceCollection services)
        where T : TeamsBotApplication
    {
        services.AddHttpContextAccessor();

        // DelegatingHandler that acquires Bearer tokens via Agent SDK's IConnections.
        services.AddTransient<AgentSdkAuthHandler>();

        // Named HttpClient with the auth handler in its pipeline.
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

        // Expose the bot under its base type so TeamsSdkMiddleware — which
        // depends on TeamsBotApplication rather than any concrete subclass — resolves
        // the same singleton instance.
        services.AddSingleton<TeamsBotApplication>(sp => sp.GetRequiredService<T>());

        return services;
    }
}

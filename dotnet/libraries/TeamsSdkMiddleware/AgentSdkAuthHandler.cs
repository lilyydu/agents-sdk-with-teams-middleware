// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Agents.Authentication;

namespace TeamsSdk;

/// <summary>
/// A <see cref="DelegatingHandler"/> that bridges the Agent SDK's authentication
/// (<see cref="IConnections"/> / <see cref="IAccessTokenProvider"/>) into the
/// Teams SDK's outbound HTTP pipeline, so the Teams SDK's <c>ConversationClient</c>
/// sends authenticated requests without a separate AzureAd config section.
/// </summary>
internal class AgentSdkAuthHandler(
    IConnections connections) : DelegatingHandler
{
    private readonly IConnections _connections = connections;

    /// <inheritdoc/>
    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        IAccessTokenProvider tokenProvider = GetTokenProvider(request.RequestUri!);

        if (tokenProvider != null)
        {
            string serviceUrl = request.RequestUri!.GetLeftPart(System.UriPartial.Authority);
            var scopes = tokenProvider.ConnectionSettings.Scopes;
            string token = await tokenProvider.GetAccessTokenAsync(serviceUrl, scopes).ConfigureAwait(false);

            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        }

        return await base.SendAsync(request, cancellationToken).ConfigureAwait(false);
    }

    private IAccessTokenProvider GetTokenProvider(System.Uri requestUri)
    {
        string serviceUrl = requestUri.GetLeftPart(System.UriPartial.Authority);
        ClaimsIdentity? claimsIdentity = TeamsSdkMiddleware.CurrentTurnContext?.Identity;

        if (claimsIdentity?.IsAuthenticated == true)
        {
            var provider = _connections.GetTokenProvider(claimsIdentity, serviceUrl);
            if (provider != null)
            {
                return provider;
            }
        }

        return _connections.GetDefaultConnection();
    }
}

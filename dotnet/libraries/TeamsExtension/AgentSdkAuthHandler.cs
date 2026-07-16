// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Agents.Authentication;
using Microsoft.AspNetCore.Http;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Threading;
using System.Threading.Tasks;

namespace TeamsExtension;

/// <summary>
/// A <see cref="DelegatingHandler"/> that bridges Agent SDK auth
/// (<see cref="IConnections"/> / <see cref="IAccessTokenProvider"/>) into the
/// Teams SDK outbound HTTP pipeline.
/// </summary>
internal class AgentSdkAuthHandler(
    IConnections connections,
    IHttpContextAccessor httpContextAccessor) : DelegatingHandler
{
    private readonly IConnections _connections = connections;
    private readonly IHttpContextAccessor _httpContextAccessor = httpContextAccessor;

    /// <inheritdoc/>
    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        IAccessTokenProvider tokenProvider = GetTokenProvider(request.RequestUri!);

        if (tokenProvider != null)
        {
            string serviceUrl = request.RequestUri!.GetLeftPart(UriPartial.Authority);
            var scopes = tokenProvider.ConnectionSettings.Scopes;
            string token = await tokenProvider.GetAccessTokenAsync(serviceUrl, scopes).ConfigureAwait(false);

            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        }

        return await base.SendAsync(request, cancellationToken).ConfigureAwait(false);
    }

    private IAccessTokenProvider GetTokenProvider(Uri requestUri)
    {
        string serviceUrl = requestUri.GetLeftPart(UriPartial.Authority);
        var claimsIdentity = _httpContextAccessor.HttpContext?.User?.Identity as ClaimsIdentity;

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

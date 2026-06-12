/**
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 *
 * Adapt Agents SDK's connection manager (MsalConnectionManager) to teams.ts's
 * token-callback shape.
 *
 * teams.ts's AppOptions accepts a `token` callable of shape
 * `(scope, tenantId) => string | Promise<string>` which is wrapped into
 * TokenCredentials internally. Agents SDK already has a working MSAL provider;
 * this module produces the callback so developers don't have to duplicate
 * token logic.
 *
 * Connection-aware:
 *   When invoked inside a turn handled by TeamsSdkMiddleware, the callback
 *   reads the inbound identity from `agentSdkTurnContext().identity` and asks
 *   the connection manager for the matching connection via
 *   `getTokenProvider(identity, serviceUrl)`. This matters when more than one
 *   connection is registered (e.g. different app registrations per channel
 *   or skill audience); with a single connection it returns the same provider
 *   as `getDefaultConnection()`.
 *
 *   Note: this does NOT swap tenants per user. Outbound Bot Framework Service
 *   tokens are always acquired against the bot's home tenant configured on
 *   the connection itself; the inbound identity only selects WHICH connection
 *   to use, not the tenant inside that connection.
 *
 *   Outside a turn (proactive callbacks, background tasks, startup hooks) the
 *   AsyncLocalStorage is unset and we fall back to `getDefaultConnection()`.
 */

import type { AuthProvider } from '@microsoft/agents-hosting';

import { _agentSdkTurnContextStore } from './context';
import type { AgentSdkConnections } from './install';

const DEFAULT_SUFFIX = '/.default';

export type TeamsSdkTokenCallback = (
  scope: string | string[],
  tenantId?: string
) => Promise<string>;

/**
 * Return a token callback that delegates to a connection manager.
 *
 * @param connectionManager Agents SDK connection manager (typically
 *   MsalConnectionManager). Must expose `getDefaultConnection()` and
 *   `getTokenProvider(identity, serviceUrl)`, each returning an AuthProvider
 *   with `getAccessToken(scope)`.
 *
 * @returns An async callable matching teams.ts's TokenCredentials.token
 *   signature. Connection-aware (picks among multiple registered
 *   connections when applicable) when invoked inside a turn handled by
 *   TeamsSdkMiddleware; falls back to the default connection for proactive /
 *   background calls.
 */
export function createAgentSdkTokenProvider(
  connectionManager: AgentSdkConnections
): TeamsSdkTokenCallback {
  return async (scope, _tenantId) => {
    const scopes = Array.isArray(scope) ? scope : [scope];
    // Strip "/.default" off the first scope to derive the resource_url MSAL
    // wants — '/.default' is appended back internally.
    const first = scopes[0] ?? '';
    const resourceUrl = first.endsWith(DEFAULT_SUFFIX)
      ? first.slice(0, -DEFAULT_SUFFIX.length)
      : first;

    const provider = selectProvider(connectionManager, resourceUrl);
    return provider.getAccessToken(resourceUrl);
  };
}

/**
 * Pick the matching connection when a turn is in scope, else default.
 *
 * Looks up the inbound identity from the current turn's `context.identity`. If
 * anything is missing — no turn in scope, no identity stashed, or the
 * connection manager rejects the lookup — fall back to the default connection.
 * That fallback is the only path used by proactive sends and non-turn callers.
 *
 * NOTE: This does NOT perform per-user-tenant token acquisition. The chosen
 * connection's own tenantId is always used for the outbound token. This call
 * only matters when more than one connection is registered.
 */
function selectProvider(
  connectionManager: AgentSdkConnections,
  serviceUrl: string
): AuthProvider {
  const context = _agentSdkTurnContextStore.getStore();
  if (!context) {
    return connectionManager.getDefaultConnection();
  }

  // Property access can throw if the TurnContext proxy has been revoked
  // (e.g. delayed proactive sends from setTimeout/setInterval callbacks that
  // outlive the originating turn — AsyncLocalStorage still propagates the
  // reference, but agents-hosting has already torn the proxy down).
  try {
    const identity = context.identity;
    if (!identity) {
      return connectionManager.getDefaultConnection();
    }

    const targetUrl = serviceUrl || context.activity.serviceUrl || '';
    if (!targetUrl) {
      return connectionManager.getDefaultConnection();
    }

    return connectionManager.getTokenProvider(identity, targetUrl);
  } catch {
    // Degrade gracefully on revoked-proxy access or lookup failure.
    return connectionManager.getDefaultConnection();
  }
}

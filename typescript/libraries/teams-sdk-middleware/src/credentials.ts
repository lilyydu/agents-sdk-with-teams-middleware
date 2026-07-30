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
 *
 * Agentic (Agent 365) requests:
 *   When teams.ts needs an outbound token to act as the *agentic user*
 *   (Teams-channel agentic turns), it invokes this callback with a third
 *   argument carrying an `agenticIdentity`. In that case we mint the token via
 *   the connection's `getAgenticUserToken` — the same Agents SDK path as
 *   AgenticUserAuthorization — instead of the normal bot token. The bot's own
 *   client_id is the blueprint the token derives from; teams.ts supplies the
 *   agentic app-instance id, user id, and scopes. Agentic *bot* (app-instance)
 *   outbound is intentionally not handled here: teams.ts has no such path, so
 *   use `agentSdkTurnContext().sendActivity` for that case.
 */

import type { AuthProvider } from '@microsoft/agents-hosting';

import { _agentSdkTurnContextStore } from './context';
import type { AgentSdkConnections } from './install';

const DEFAULT_SUFFIX = '/.default';

/**
 * Routing key used to select the connection that mints agentic tokens,
 * mirroring the Agents SDK's `getTokenProvider(identity, 'agentic')` convention.
 * With a single registered connection the lookup falls back to the default
 * connection, which is sufficient for the common case.
 */
const AGENTIC_ROUTING_KEY = 'agentic';

/**
 * The agentic identity teams.ts passes to the token callback for agentic-user
 * turns. Declared locally because the published `@microsoft/teams.api` barrel
 * does not yet export `AgenticIdentity` (it ships on the Agent 365 line); this
 * structural shape matches what teams.ts provides.
 */
export interface AgenticIdentity {
  readonly agenticAppId: string;
  readonly agenticUserId: string;
  readonly tenantId?: string;
  readonly agenticAppBlueprintId?: string;
}

/** Third-argument options bag teams.ts passes to the token callback. */
export interface TokenRequestOptions {
  readonly agenticIdentity?: AgenticIdentity;
}

/**
 * Agentic-capable Agents SDK provider. The published `AuthProvider` type does
 * not yet declare `getAgenticUserToken`, so we extend it structurally and probe
 * at runtime — present on an agentic-capable `@microsoft/agents-hosting`.
 */
type AgenticCapableProvider = AuthProvider & {
  getAgenticUserToken?: (
    tenantId: string,
    agentAppInstanceId: string,
    upn: string,
    scopes: string[]
  ) => Promise<string>;
};

export type TeamsSdkTokenCallback = (
  scope: string | string[],
  tenantId?: string,
  options?: TokenRequestOptions
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
  return async (scope, tenantId, options) => {
    const scopes = Array.isArray(scope) ? scope : [scope];

    // Agentic-user path: teams.ts passes an `agenticIdentity` when the outbound
    // call must act as the agentic user rather than the bot. Mint the user
    // token through the same Agents SDK connection (single source of truth),
    // mirroring AgenticUserAuthorization. teams.ts only populates this for
    // agentic turns; otherwise we fall through to the normal bot token.
    const agenticIdentity = options?.agenticIdentity;
    if (agenticIdentity?.agenticUserId) {
      const provider = selectProvider(
        connectionManager,
        AGENTIC_ROUTING_KEY
      ) as AgenticCapableProvider;
      if (typeof provider.getAgenticUserToken !== 'function') {
        throw new Error(
          'Agents SDK connection does not expose getAgenticUserToken; an ' +
          'agentic-capable @microsoft/agents-hosting (MSAL provider) is required.'
        );
      }
      return provider.getAgenticUserToken(
        agenticIdentity.tenantId ?? tenantId ?? '',
        agenticIdentity.agenticAppId,
        agenticIdentity.agenticUserId,
        scopes
      );
    }

    // Non-agentic path: outbound Bot Framework Service token. Strip "/.default"
    // off the first scope to derive the resource_url MSAL wants — '/.default'
    // is appended back internally.
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

/**
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 *
 * One-call factory: build a teams.ts App from an Agents SDK app + connections.
 *
 * This is the recommended entry point. It owns the three things that otherwise
 * have to be wired by hand at every call site:
 *
 *   1. Credential extraction — pulls clientId/tenantId from the connection
 *      manager's default configuration.
 *   2. Token-provider wiring — bridges teams.ts's outbound token callback to
 *      the Agents SDK MsalConnectionManager.
 *   3. Middleware install — registers TeamsSdkMiddleware on the Agents SDK
 *      adapter so Teams turns are short-circuited to teams.ts.
 *
 * Returns a fully constructed teams.ts App ready for handler registration.
 */

import type {
  AgentApplication,
  AuthConfiguration,
  AuthProvider,
  TurnState,
} from '@microsoft/agents-hosting';
import { App, type AppOptions, type IPlugin } from '@microsoft/teams.apps';

import { createAgentSdkTokenProvider } from './credentials';
import { TeamsSdkMiddleware } from './middleware';

const RESERVED_KEYS = ['clientId', 'tenantId', 'token'] as const;

/**
 * The subset of @microsoft/agents-hosting's `Connections` interface we depend on.
 *
 * Defined locally because the published `@microsoft/agents-hosting` barrel does
 * not currently re-export the `Connections` type. `MsalConnectionManager` (and
 * any other Connections implementation) satisfies this structurally.
 *
 * `identity` is typed as `unknown` to avoid pulling in @types/jsonwebtoken as
 * a dependency — pass through whatever `TurnContext.identity` returns.
 */
export interface AgentSdkConnections {
  getDefaultConnection(): AuthProvider;
  getTokenProvider(identity: unknown, serviceUrl: string): AuthProvider;
  getDefaultConnectionConfiguration(): AuthConfiguration;
}

/**
 * Options forwarded to the teams.ts App constructor, minus the credential
 * fields that `useTeamsSdk` owns.
 */
export type UseTeamsSdkOptions<TPlugin extends IPlugin = IPlugin> = Omit<
  AppOptions<TPlugin>,
  (typeof RESERVED_KEYS)[number]
>;

/**
 * Wire teams.ts into `app` and return the configured App.
 *
 * @param app The Agents SDK application whose adapter will own the HTTP
 *   endpoint. TeamsSdkMiddleware is installed on `app.adapter`.
 * @param connectionManager Source of credentials and tokens. The default
 *   connection's clientId and tenantId are used to construct the teams.ts
 *   App; its token providers are wrapped so teams.ts's outbound calls use
 *   the same credentials.
 * @param teamsAppOptions Extra options forwarded to the teams.ts App
 *   constructor. Use this for `logger`, `plugins`, or any other AppOptions
 *   field. `clientId`, `tenantId`, and `token` are reserved and will throw
 *   if passed here.
 *
 * @returns The configured teams.ts App. Register handlers on the returned
 *   object (`teamsApp.on('message', ...)`, etc.).
 *
 * After this call:
 *   - Teams turns with a matching teams.ts handler → handled by the App.
 *   - Teams turns with no match → fall through to `app`'s handlers.
 *   - Any other channel → handled by `app` unchanged.
 */
export function useTeamsSdk<TState extends TurnState, TPlugin extends IPlugin = IPlugin>(
  app: AgentApplication<TState>,
  connectionManager: AgentSdkConnections,
  teamsAppOptions: UseTeamsSdkOptions<TPlugin> = {}
): App<TPlugin> {
  const reserved = RESERVED_KEYS.filter(k => k in teamsAppOptions);
  if (reserved.length > 0) {
    throw new TypeError(
      `useTeamsSdk owns ${JSON.stringify(reserved)}; remove from teamsAppOptions.`
    );
  }

  const auth = connectionManager.getDefaultConnectionConfiguration();
  const teamsApp = new App<TPlugin>({
    ...(teamsAppOptions as AppOptions<TPlugin>),
    clientId: auth.clientId,
    tenantId: auth.tenantId,
    token: createAgentSdkTokenProvider(connectionManager),
  });

  app.adapter.use(new TeamsSdkMiddleware(teamsApp));
  return teamsApp;
}

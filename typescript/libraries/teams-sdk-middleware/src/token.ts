/**
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 *
 * Synthesize a teams.ts IToken from an inbound Agents SDK Activity.
 *
 * teams.ts's IActivityEvent requires an IToken so the framework can build its
 * context (service_url fallback, caller classification, app_id). The Agents
 * SDK has already validated the inbound JWT by the time the middleware runs,
 * so we don't re-forward the raw bearer — we just project the fields teams.ts
 * actually consumes off the Activity itself.
 *
 * Outbound auth goes through teams.ts's TokenCredentials, which the install
 * factory sets up separately via createAgentSdkTokenProvider.
 */

import type { Activity } from '@microsoft/agents-activity';
import type { IToken } from '@microsoft/teams.api';

const ONE_HOUR_MS = 60 * 60 * 1000;
const DEFAULT_BUFFER_MS = 5 * 60 * 1000;
const DEFAULT_SERVICE_URL = 'https://smba.trafficmanager.net/teams';

// teams.ts CallerIds mirror — kept local to avoid an extra import.
const CALLER_ID_BOT = 'urn:botframework:aadappid';
const CALLER_ID_AZURE = 'urn:botframework:azure';

/**
 * A minimal IToken synthesized from an Agents SDK Activity.
 *
 * Carries no signable bearer — Agents SDK has already authenticated the
 * inbound request, and outbound auth flows through TokenCredentials on the
 * teams.ts App.
 */
export class TeamsSdkSyntheticToken implements IToken {
  readonly appId: string;
  readonly appDisplayName?: string;
  readonly tenantId?: string;
  readonly serviceUrl: string;
  readonly from: 'bot' | 'azure';
  readonly fromId: string;
  readonly expiration: number;

  constructor(options: {
    appId: string;
    appDisplayName?: string;
    tenantId?: string;
    serviceUrl: string;
  }) {
    this.appId = options.appId;
    this.appDisplayName = options.appDisplayName;
    this.tenantId = options.tenantId;
    const trimmed = (options.serviceUrl ?? '').replace(/\/+$/, '');
    this.serviceUrl = trimmed || DEFAULT_SERVICE_URL;
    this.from = this.appId ? 'bot' : 'azure';
    this.fromId = this.from === 'bot'
      ? `${CALLER_ID_BOT}:${this.appId}`
      : CALLER_ID_AZURE;
    // Synthetic expiration ~1h out so isExpired() reports false.
    this.expiration = Date.now() + ONE_HOUR_MS;
  }

  static fromActivity(activity: Activity): TeamsSdkSyntheticToken {
    return new TeamsSdkSyntheticToken({
      appId: activity.recipient?.id ?? '',
      appDisplayName: activity.recipient?.name,
      tenantId: activity.conversation?.tenantId,
      serviceUrl: activity.serviceUrl ?? '',
    });
  }

  isExpired(bufferMs: number = DEFAULT_BUFFER_MS): boolean {
    return this.expiration < Date.now() + bufferMs;
  }

  toString(): string {
    // No real JWT is forwarded; a tagged sentinel is more honest than '' and
    // cannot be mistaken for a signable bearer.
    return `teams-sdk-synthetic://app/${this.appId}`;
  }
}

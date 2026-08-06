/**
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 *
 * Middleware that matches Teams turns to a teams.ts App.
 *
 * Lifecycle of a turn:
 *
 *   1. Non-Teams channel → pass through to AgentApplication.
 *   2. Teams turn the caller's `shouldBypassTeams` predicate claims → pass
 *      through too, even when teams.ts has a matching route.
 *   3. Teams turn with no matching teams.ts route → pass through too.
 *   4. Teams turn with a match → ensure teams.ts is initialized, expose the
 *      Agents SDK TurnContext via AsyncLocalStorage, hand the activity to
 *      teams.ts's process(), then propagate any InvokeResponse back through
 *      the Agents SDK send pipeline so the HTTP layer can write the
 *      synchronous body.
 *
 * The middleware never calls `next()` after teams.ts has handled a turn —
 * doing so would invoke AgentApplication's handlers a second time for the
 * same activity and, for invokes, would emit a duplicate response.
 */

import {
  type Activity as AgentsActivity,
  ActivityTypes,
} from '@microsoft/agents-activity';
import {
  type Middleware,
  type TurnContext,
} from '@microsoft/agents-hosting';
import type { Activity as TeamsActivity, InvokeResponse } from '@microsoft/teams.api';
import type { App } from '@microsoft/teams.apps';

import { _agentSdkTurnContextStore } from './context';
import { TeamsSdkSyntheticToken } from './token';

const TEAMS_CHANNEL_ID = 'msteams';

/**
 * Predicate evaluated only for Teams-channel turns. Return `true` to bypass
 * teams.ts routing and force the turn to fall through to AgentApplication even
 * when teams.ts has a matching route.
 */
export type ShouldBypassTeams = (context: TurnContext) => boolean;

/** True for Teams turns, including sub-channels like `msteams:COPILOT`. */
export function isTeamsChannel(activity: { channelId?: string }): boolean {
  const channelId = activity.channelId;
  if (!channelId) {
    return false;
  }
  return channelId.split(':', 1)[0] === TEAMS_CHANNEL_ID;
}

export class TeamsSdkMiddleware implements Middleware {
  private readonly _teamsApp: App<any>;
  private readonly _shouldBypassTeams?: ShouldBypassTeams;

  /**
   * @param teamsApp The teams.ts App that owns matching Teams turns.
   * @param shouldBypassTeams Optional predicate that claims Teams turns for
   *   AgentApplication.
   */
  constructor(teamsApp: App<any>, shouldBypassTeams?: ShouldBypassTeams) {
    this._teamsApp = teamsApp;
    this._shouldBypassTeams = shouldBypassTeams;
  }

  async onTurn(context: TurnContext, next: () => Promise<void>): Promise<void> {
    if (!isTeamsChannel(context.activity)) {
      await next();
      return;
    }

    // Idempotent — App tracks isInitialized internally. Run on every Teams
    // turn so AgentApplication handlers can safely call into TEAMS_APP (e.g.
    // TEAMS_APP.send for proactive sends) even when no teams.ts route matched
    // this turn.
    await this._teamsApp.initialize();

    // teams.ts registers signin/tokenExchange, signin/verifyState and
    // signin/failure handlers unconditionally at App construction, so the
    // router matches them even for a flow AgentApplication started. Callers
    // that keep authorization on the Agents SDK side claim those turns here,
    // before the activity is handed to teams.ts at all.
    if (this._shouldBypassTeams?.(context)) {
      await next();
      return;
    }

    const coreActivity = context.activity as unknown as TeamsActivity;

    // Router is protected on App; reach for it once and treat absence as
    // "no routes matched" so we degrade to the fallthrough path.
    const router: any = (this._teamsApp as any).router;
    const matched = router?.select?.(coreActivity);
    if (!matched || matched.length === 0) {
      // No teams.ts route matches; let AgentApplication try its handlers.
      await next();
      return;
    }

    const event = {
      body: coreActivity,
      token: TeamsSdkSyntheticToken.fromActivity(context.activity),
    };

    let invokeResponse: InvokeResponse | undefined;
    await _agentSdkTurnContextStore.run(context, async () => {
      invokeResponse = await this._teamsApp.process(event);
    });

    if (context.activity.type === ActivityTypes.Invoke) {
      await this._propagateInvokeResponse(context, invokeResponse);
    }
  }

  /**
   * Emit teams.ts's InvokeResponse through the Agents SDK send pipeline.
   *
   * Wraps the response in an outbound activity of type `invokeResponse`;
   * CloudAdapter intercepts it via sendActivities and stashes it under
   * INVOKE_RESPONSE_KEY so the HTTP layer writes it as the synchronous
   * response body.
   */
  private async _propagateInvokeResponse(
    context: TurnContext,
    invokeResponse: InvokeResponse | undefined
  ): Promise<void> {
    if (!invokeResponse || invokeResponse.status === undefined) {
      return;
    }

    const out = {
      type: ActivityTypes.InvokeResponse,
      value: {
        status: invokeResponse.status,
        body: invokeResponse.body,
      },
    } as unknown as AgentsActivity;

    await context.sendActivity(out);
  }
}

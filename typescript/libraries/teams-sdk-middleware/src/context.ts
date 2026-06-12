/**
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 *
 * Async-task-local access to the Agents SDK TurnContext for the in-flight turn.
 *
 * Equivalent to Python's ContextVar — uses Node's AsyncLocalStorage so any code
 * running inside the async call tree of TeamsSdkMiddleware.onTurn can pull the
 * upstream TurnContext (e.g., from a teams.ts handler that wants to read/write
 * turn_state or call sendActivity on the Agents SDK pipeline).
 */

import { AsyncLocalStorage } from 'node:async_hooks';
import type { TurnContext } from '@microsoft/agents-hosting';

/** @internal Storage handle — set/reset by TeamsSdkMiddleware. */
export const _agentSdkTurnContextStore = new AsyncLocalStorage<TurnContext>();

/**
 * Return the Agents SDK TurnContext for the current turn.
 *
 * @throws Error if called outside a turn handled by TeamsSdkMiddleware (e.g.,
 *   from a startup hook, a background task that doesn't inherit the turn's
 *   async context, or a non-Teams turn that fell through).
 */
export function agentSdkTurnContext(): TurnContext {
  const ctx = _agentSdkTurnContextStore.getStore();
  if (!ctx) {
    throw new Error(
      'agentSdkTurnContext() called outside a Teams SDK middleware turn. ' +
      'This usually means the caller is running on a detached async task ' +
      '(setTimeout, queueMicrotask without inheriting context, etc.) or ' +
      'outside the request lifecycle entirely.'
    );
  }
  return ctx;
}

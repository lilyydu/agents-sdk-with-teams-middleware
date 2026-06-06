# Teams Extension in `microsoft/Agents-for-python` — Companion RFC

> Author: Lily Du · 2026-06-05
> Status: Draft
> Companion to: [aamir-architecture-diagrams.md](https://github.com/rajan-chari/fellow-scholars/blob/main/RFC/agent_sdk_interop/aamir-architecture-diagrams.md)

## Why this exists

The Teams SDK team needs to own the rich Teams interaction surface inside the
Agents SDK without forking it. Aamir's RFC proposes the right shape; this doc
locks down the contract, fills two gaps, and provides a working skeleton.

## TL;DR

Adopt Aamir's two-axis split (ChannelAdapter for plumbing, DeepExtension for
surface) and his progressive-adoption phases. Add three things he doesn't cover:

1. **Decorator cadence** — escape-hatch decorators (`on_invoke(name)`,
   `on_activity(filter)`) so devs are never blocked when Teams ships a new
   activity type. Curated typed decorators are sugar on top.
2. **Versioning contract** — `name / version / requires_core` on every
   extension; enforced at `app.use(...)`.
3. **Owner-tagged routes / ExtensionRegistry** — explicit ownership of routes
   and channels rather than fragile priority numbers across extensions.

No code is moved out of core. Channel-agnostic concerns (Authorization,
Proactive, RestChannelServiceClientFactory) stay in core. The Teams Extension
**registers into** those existing systems at `initialize()`.

## The contract (8 methods total)

```python
class ChannelAdapter(Protocol):
    name: str; version: str; requires_core: str
    def owns_channel(self, channel_id: str) -> bool: ...
    async def on_context_created(self, ctx: TurnContext) -> None: ...
    async def on_activity_sending(self, ctx, activity) -> Activity: ...

class DeepExtension(Protocol):
    name: str; version: str; requires_core: str
    def owns_channel(self, channel_id: str) -> bool: ...
    def can_handle(self, ctx: TurnContext) -> bool: ...
    async def handle(self, ctx: TurnContext) -> None: ...
```

A package can ship one, the other, or both. `app.use(...)` accepts either.

## What changes in core (small, additive)

Three additions, zero removals:

1. **`core/extension.py`** — the two Protocols + an `Extension` bundle.
2. **`core/extension_registry.py`** — `register_*`, `channel_adapter_for`,
   `deep_extension_for`.
3. **Three new call sites** in existing core files (annotated in
   `src/microsoft_agents/hosting/core/_diffs.py`):
   - `CloudAdapter.process()` — call `adapter.on_context_created(ctx)` after
     TurnContext is built.
   - `CloudAdapter.send_activities()` (or equivalent) — call
     `adapter.on_activity_sending(ctx, act)` before egress.
   - `Proactive.continue_conversation()` — call `on_context_created` after
     building the proactive TurnContext.

`RestChannelServiceClientFactory` is **unchanged**. The Teams ChannelAdapter
swaps the connector client on the TurnContext in `on_context_created`. Core's
default factory still produces the generic client.

## What we adopt from Aamir

- Two-axis split (plumbing vs surface).
- `on_activity_sending` outbound hook (we missed this; it's important for
  mentions, channelData, card transforms).
- Single `applyChannelAdapter` semantics invoked from both reactive and
  proactive entry points.
- Progressive adoption phases 1 → 4 (see `examples/`).

## Where we differ from Aamir

| Concern | Aamir | This RFC |
|---|---|---|
| "Teams wins" mechanism | Route priorities (0/1/3) | Owner-tagged routes via `ExtensionRegistry`; priorities still allowed within an owner |
| Decorator cadence | Not addressed | Escape hatch + curated sugar |
| Versioning | Not addressed | `requires_core` SemVer at registration |
| Multiple channel extensions | Implicit (Teams-only diagrams) | `ExtensionRegistry` resolves Slack/Outlook/Teams co-existence |

## Rich Teams handler surface — decorator strategy

```python
# Layer 1 — escape hatch, NEVER lags Teams releases
@teams.on_invoke("composeExtension/queryNewThing")
async def h(ctx, payload):  # payload = raw dict
    return result

@teams.on_activity(lambda a: a.type == "messageUpdate"
                          and (a.channel_data or {}).get("eventType") == "editMessage")
async def edited(ctx):
    ...

# Layer 2 — curated typed sugar (one line per typed wrapper)
@teams.message_extension.on_query
async def search(ctx, query: MessagingExtensionQuery):
    ...
```

Adding a new typed wrapper is one line in a manifest. Adding support for a new
Teams invoke type **is not required** to use it — Layer 1 always works.

## Phases (see `examples/`)

| Phase | File | What you wrote |
|---|---|---|
| 1 — Agents SDK only | `examples/phase1_basic.py` | `@app.on_message` |
| 2 — + Plumbing | `examples/phase2_plumbing.py` | `app.use(TeamsChannelAdapter(...))` |
| 3 — + Deep surface | `examples/phase3_deep.py` | `app.use(TeamsDeepExtension(...))` + Teams decorators |
| 4 — + Override | `examples/phase4_override.py` | `@teams.on_message` (Teams handles text on Teams channel only) |

## Open questions

1. **TeamsContext drift.** `TeamsContext.send(...)` and `ctx.send_activity(...)`
   must both funnel through `on_activity_sending`. Verify in implementation.
2. **Cross-extension precedence.** When Slack + Teams DeepExtensions both
   register, registry returns by channel_id — but what if an activity has no
   channel_id? Define a deterministic tiebreak (registration order is fine).
3. **Phase-4 override semantics.** Replace vs additive when
   `teams.on_message` is registered? **Recommendation: replace** — that's what
   "override" means and what users expect.
4. **Auth ownership.** Authorization stays in core; Teams-specific
   `AuthHandler`s (e.g., Teams SSO) are *registered* by the extension via
   `host.authorization.register_handler(...)`. Confirm scope.

## Implementation checklist (smallest credible first PR)

- [ ] `core/extension.py` — protocols
- [ ] `core/extension_registry.py` — registry
- [ ] `AgentApplication.use(extension)` — register helper
- [ ] Three call sites: `process()`, `send_activities()`,
      `continue_conversation()`
- [ ] `TeamsChannelAdapter` — implements `ChannelAdapter`
- [ ] `TeamsDeepExtension` — wraps existing `TeamsAgentExtension`'s decorators,
      adds `on_invoke` / `on_activity` escape hatch
- [ ] `TeamsContext` — bridge from TurnContext (start with `{send, stream,
      graph, signin}`; expand over time)
- [ ] Example apps for phases 1–4
- [ ] One end-to-end test per phase

That's ~7 files for a real PR. Skeleton in `src/` and `examples/`.

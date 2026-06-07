# RFC: Teams Extension for the M365 Agents SDK (Python)

**Author:** Teams SDK Team
**Status:** Draft
**Repo:** `microsoft/Agents-for-python`
**Sample code:** `test_samples/teams_extension_rfc/`

---

## TL;DR

Teams support for the M365 Agents SDK for Python ships as a **standalone
middleware + decorator surface** in `microsoft-agents-hosting-teams`, with
**zero changes to `microsoft-agents-hosting-core`**.

The Teams team owns the entire Teams developer experience — plumbing,
invokes, mentions, cards, notifications, sign-in — from inside the Teams
package only.

---

## Goals (recap)

1. Teams SDK team owns the rich Teams interaction surface in any SDK.
2. Teams SDK team owns the client that connects to Teams.
3. Rich Teams experience must not be compromised by lowest-common-denominator
   chat features (and vice versa).
4. Devs can move between standalone Teams SDK and Agents SDK with the same
   API.
5. Teams SDK team owns release cadence for the Teams Extension.
6. Frictionless to build and deploy Teams agents using any SDK.

## Non-goals (this RFC)

* Replacing the existing `AgentApplication` programming model.
* Adding new extension points to `microsoft-agents-hosting-core`.
* Building a formal "Extension Protocol" abstraction in core.
* Supporting third-party channel teams in v1 — the pattern below is
  copy-able for them, but no formal contract is shipped.

---

## Design

### Two components, both in `microsoft-agents-hosting-teams`

1. **`TeamsMiddleware`** — *plumbing*. One instance, registered on the
   adapter. Runs on every turn (reactive, proactive, continue-conversation
   — all flows go through `ChannelServiceAdapter.run_pipeline`). Filters
   internally on `channel_id == "msteams"`.

   Inbound work (replaces the would-be `on_context_created`):
   * Swap the `ConnectorClient` for `TeamsConnectorClient`.
   * Attach Teams helpers (`TeamsInfo`, parsed channelData).

   Outbound work (replaces the would-be `on_activity_sending`):
   * Registers a callback with the existing
     `TurnContext.on_send_activities(handler)` API.
   * Injects notification metadata.
   * Rewrites mentions into Teams `<at>` form.
   * Transforms Adaptive Cards for Teams quirks.

2. **`TeamsHandlers`** — *surface*. Thin wrapper around the existing
   `AgentApplication.add_route(selector, handler, is_invoke=True, rank)`
   API. Two layers:

   * **Layer 1 — escape hatch.** `on_invoke(name)`, `on_activity(predicate)`.
     Works day-1 for any new Teams invoke type, without an SDK release.
   * **Layer 2 — curated sugar.** Typed, ergonomic decorators for the common
     invoke types: `message_extension_query`, `task_module_fetch`,
     `task_module_submit`, `adaptive_card_action_execute`,
     `sign_in_verify_state`, …

### Developer experience

```python
from microsoft_agents.hosting.core import AgentApplication
from microsoft_agents.hosting.teams import install_teams, TeamsHandlers

app = AgentApplication()
install_teams(app)                  # plumbing
teams = TeamsHandlers(app)          # surface

@app.message("hello")
async def hello(ctx, state):
    await ctx.send_activity("Hi there!")   # Teams-correct on the wire

@teams.message_extension_query("search")
async def search(ctx, query):
    return MessagingExtensionResponse(...)
```

### What we lean on in core (already exists)

| API | Used for |
| --- | --- |
| `CloudAdapter.use(middleware)` | Registers `TeamsMiddleware` so it runs on every turn (all entry points). |
| `Middleware.on_turn` | Inbound enrichment hook. |
| `TurnContext.on_send_activities(handler)` | Outbound transform hook. Already part of `TurnContext`, lines 293-300. |
| `AgentApplication.add_route(selector, handler, is_invoke=True, rank, auth_handlers)` | Route registration. Already part of `AgentApplication`, line 248. |
| `RouteRank` | Invokes ranked higher than messages — already enforced by the dispatcher. |
| `turn_state["BotFrameworkAdapter.InvokeResponse"]` | How handlers return typed invoke bodies — existing convention. |

---

## Why not formalise `ChannelAdapter` / `DeepExtension` Protocols in core?

An earlier draft of this RFC proposed adding two Protocols and a registry
to core. We rejected that approach for v1:

* **Scope creep into a package the Teams team doesn't own.** Even a 30-line
  change to core needs core-team review, regression testing, and release
  coordination.
* **YAGNI.** Today there is one extension (Teams). The cost-benefit only
  flips when a second channel team needs the same pattern.
* **The existing primitives already do everything we need.**
  `Middleware.on_turn` + `TurnContext.on_send_activities` + `add_route`
  cover plumbing, outbound transforms, and routing respectively.

When a second channel team needs the same shape, the existing
`TeamsMiddleware` + `TeamsHandlers` pattern is the spec. Promoting it to a
formal Protocol/registry in core becomes a mechanical refactor, not a
ground-up design.

---

## What this design *cannot* do (and why that's OK)

| Limitation | Mitigation |
| --- | --- |
| No `requires_core` SemVer enforcement at startup. | The Teams package pins a core version in its `pyproject.toml` as usual. Mismatch surfaces at `pip install` time. |
| No introspection of "what extensions are loaded". | Middleware list on the adapter is the answer. Diagnostics tooling can read it. |
| No structural Protocol for other channel teams to conform to. | They copy the `TeamsMiddleware` shape; the pattern is documented in this RFC and in the package README. |
| `Middleware.on_turn` runs on every turn even on non-Teams channels. | Single `if channel_id == "msteams"` short-circuit. Cost: one branch per turn. |

---

## Migration

There is no migration for end-users. New code looks like the sample.
Existing Teams code using `TeamsActivityHandler` continues to work; that
class is unaffected by this RFC.

---

## Open questions

1. **Where does the `TeamsConnectorClient` get its credentials?** The
   middleware needs access to whatever the adapter used to build the
   default connector. Either expose a credentials provider on the adapter,
   or read it off the existing client and reuse.

2. **Auth handler integration.** `add_route` accepts `auth_handlers`. The
   curated decorators forward them. Does that satisfy the goal of "SSO/OAuth
   lives in the extension", or do we need a Teams-specific auth handler
   registration helper?

3. **Streaming.** Teams has size-chunking requirements for streaming
   responses. Does this belong in `TeamsMiddleware._on_send_activities`,
   or as a separate streaming-specific decorator?

4. **Drift risk.** If a developer ever sends activities through a path that
   bypasses `TurnContext.send_activity` (e.g. calling the connector client
   directly), the outbound transforms are skipped. Document this as a known
   limitation and a code-review smell.

---

## Implementation checklist

- [ ] `teams_middleware.py` — fill in `_swap_connector_client`,
      `_format_mentions`, `_transform_card`.
- [ ] `teams_handlers.py` — add typed payload models
      (`MessagingExtensionQuery`, `TaskModuleRequest`, …) once exported by
      `microsoft-agents-activity` or local to the package.
- [ ] Unit tests: middleware no-ops on non-Teams channels; outbound
      transforms run; invoke decorators select correctly.
- [ ] Integration test: full reactive + proactive + invoke roundtrip in a
      Teams-emulator harness.
- [ ] Migration note in package README for users of the legacy
      `TeamsAgentExtension`.

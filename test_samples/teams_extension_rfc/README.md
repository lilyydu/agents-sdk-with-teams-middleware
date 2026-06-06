# Teams Extension RFC — examples

Four files showing the adoption ramp described in `RFC-teams-extension.md`:

| File | What it adds |
|---|---|
| `phase1_basic.py` | Agents SDK only — works on every channel (LCM) |
| `phase2_plumbing.py` | + `TeamsChannelAdapter` — plumbing, no handler changes |
| `phase3_deep.py` | + `TeamsDeepExtension` — rich Teams handlers (additive) |
| `phase4_override.py` | `@teams.on_message()` — Teams handles text on Teams only |

These are *design artifacts*, not runnable apps. They import symbols that the
RFC proposes adding (e.g., `AgentApplication.use(...)`,
`microsoft_agents.hosting.teams.teams_channel_adapter`) — those don't exist on
`main` yet. They're here to make the API tangible.

Run order if you were turning this into a real PR:
1. Add `core/extension.py`, `core/extension_registry.py` (already in this branch).
2. Wire the three call sites in `channel_service_adapter.py` and
   `app/proactive/proactive.py` per `core/_diffs.py`.
3. Add `AgentApplication.use(...)` per `core/_diffs.py`.
4. Flesh out `teams_channel_adapter.py`, `teams_deep_extension.py`,
   `teams_context.py` (skeletons in this branch).
5. These four examples become end-to-end tests.

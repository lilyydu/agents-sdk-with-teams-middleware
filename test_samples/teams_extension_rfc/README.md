# Teams Extension — RFC sample

A zero-core-change design for Teams support in the Microsoft 365 Agents SDK
for Python.

## Files

| File | What it shows |
| --- | --- |
| `01_plumbing_only.py` | Generic `AgentApplication` code + one `install_teams(app)` call. Pure LCM developer experience, Teams-correct on the wire. |
| `02_invoke_escape_hatch.py` | `@teams.on_invoke("composeExtension/...")` for invoke types not yet covered by a typed decorator. Day-1 support for new Teams features. |
| `03_curated_invokes.py` | Typed sugar: message extensions, task modules, adaptive card actions. |

## Architecture

```
┌────────────────────────────────────────────────────────────────┐
│  microsoft-agents-hosting-teams  (everything Teams-specific)   │
│                                                                │
│  ┌──────────────────────┐    ┌───────────────────────────┐     │
│  │ TeamsMiddleware      │    │ TeamsHandlers             │     │
│  │ (plumbing)           │    │ (surface)                 │     │
│  │                      │    │                           │     │
│  │ on_turn()            │    │ on_invoke(name)           │     │
│  │  • swap connector    │    │ on_activity(predicate)    │     │
│  │  • attach helpers    │    │ message_extension_query() │     │
│  │  • parse channelData │    │ task_module_fetch()       │     │
│  │  • register          │    │ task_module_submit()      │     │
│  │    ctx.on_send_      │    │ adaptive_card_action_     │     │
│  │      activities()    │    │   execute()               │     │
│  └──────────────────────┘    │ sign_in_verify_state()    │     │
│                              └───────────────────────────┘     │
│                                                                │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │ install_teams(app)  → app.adapter.use(TeamsMiddleware()) │  │
│  └──────────────────────────────────────────────────────────┘  │
└────────────────────────────────────────────────────────────────┘
                              │
                              │ uses (no changes required)
                              ▼
┌────────────────────────────────────────────────────────────────┐
│  microsoft-agents-hosting-core  (UNCHANGED)                    │
│                                                                │
│  • adapter.use(middleware)                                     │
│  • ctx.on_send_activities(handler)                             │
│  • app.add_route(selector, handler, is_invoke=True, rank=...)  │
└────────────────────────────────────────────────────────────────┘
```

## Why this design

| Requirement | How it's met |
| --- | --- |
| Zero core changes | Uses `adapter.use()`, `ctx.on_send_activities()`, `app.add_route()` — all public APIs that exist today. |
| Teams team owns the surface | All code lives in `microsoft-agents-hosting-teams`. |
| Independent release cadence | Teams package ships without touching core. |
| Day-1 support for new invokes | Escape hatch (`on_invoke(name)`) works for any unknown invoke. |
| LCM chat unaffected | Middleware filters on `channel_id == "msteams"`; routes select on the same. |
| Proactive / continue_conversation covered | All flows go through `run_pipeline` — the middleware runs on each. |

# Teams SDK + Agents SDK sample (.NET)

This sample now mirrors the multichannel shape from PR #3:

- **Teams SDK owns matched Teams routes**: `help`, `react`, `quote`, `targeted`, `task`, plus reaction events.
- **Agents SDK owns everything else**: unmatched Teams turns, all non-Teams channels, `signin/*` invokes, and the `help`, `channel`, `agents sdk react`, `agents sdk proactive`, `whoami`, `mail`, `signout`, and echo routes.
- **Channel quirks are explicit**: Teams subchannels such as `msteams:COPILOT` are treated as Teams; email auth is declined up front because OAuth cards do not work there.

## Wiring

```csharp
builder.AddAgent<MyAgent>();
builder.Services.AddSingleton<IStorage, MemoryStorage>();
builder.Services.AddAgentAspNetAuthentication(builder.Configuration);
builder.Services.AddTeamsSdk<MyTeamsBot>(shouldBypassTeams: turnContext =>
    turnContext.Activity.Type == ActivityTypes.Invoke
    && !string.IsNullOrEmpty(turnContext.Activity.Name)
    && turnContext.Activity.Name.StartsWith("signin/", StringComparison.OrdinalIgnoreCase));
```

`AddTeamsSdk<MyTeamsBot>()` is the only integration call. It registers the Teams SDK bot,
bridges outbound auth through the Agents SDK connection manager plus the ambient Agents SDK
turn context, and installs the middleware that decides whether a turn stays in the Agents SDK
or is handed to the Teams SDK.

The optional bypass runs only for Teams-channel activities and can force a fallthrough to
the Agents SDK even when the Teams SDK has a matching route. This sample uses it to keep
`signin/*` invokes owned by the Agents SDK auth pipeline instead of the Teams SDK.

## Route split

| Surface | Commands / events | Owner |
|---|---|---|
| Teams messages | `help`, `react`, `quote`, `targeted`, `task` | Teams SDK |
| Teams events | message reactions | Teams SDK |
| Teams invokes | task module submit/fetch | Teams SDK |
| Teams auth invokes | `signin/*` | Agents SDK |
| Non-Teams messages | `help`, `channel`, `agents sdk react`, `agents sdk proactive`, `whoami`, `mail`, `signout`, echo | Agents SDK |

## Local config

`appsettings.json` stays checked in with placeholders. Put real credentials in one of:

- `appsettings.Development.json`
- environment variables such as `Connections__ServiceConnection__Settings__ClientId`

The minimal override file is:

```json
{
  "TokenValidation": {
    "Enabled": true,
    "Audiences": ["<client-id>"],
    "TenantId": "<tenant-id>"
  },
  "Connections": {
    "ServiceConnection": {
      "Settings": {
        "AuthType": "ClientSecret",
        "AuthorityEndpoint": "https://login.microsoftonline.com/<tenant-id>",
        "ClientId": "<client-id>",
        "ClientSecret": "<client-secret>",
        "Scopes": ["https://api.botframework.com/.default"]
      }
    }
  }
}
```

## E2E setup

### Teams bot + Azure migration + SSO

1. Create a persistent tunnel and host the sample on a stable port.

   ```powershell
   devtunnel create teams-middleware-sample --allow-anonymous
   devtunnel port create teams-middleware-sample -p 3980 --protocol auto
   devtunnel host teams-middleware-sample
   ```

2. Create the Teams-managed app, then migrate it to Azure.

   ```powershell
   teams app create --name "TeamsMiddlewareSample" --endpoint "https://<tunnel-id>-3980.<cluster>.devtunnels.ms/api/messages" --json
   teams app bot migrate <teamsAppId> --subscription <subscription-id> --resource-group <rg> --create-resource-group --region westus2 --json
   ```

3. Generate a client secret and write local runtime config.

   ```powershell
   teams app auth secret create <teamsAppId> --env .\bot.env
   ```

4. Patch the Entra app for SSO:
   - identifier URI: `api://botid-<teamsAppId>`
   - `access_as_user` scope
   - Bot Framework redirect URI
   - pre-authorized Teams desktop and web clients

5. Create the Azure Bot OAuth connections used by this sample.

   ```powershell
   az bot authsetting create -n <teamsAppId> -g <rg> -c graphuser --service Aadv2 --client-id <teamsAppId> --client-secret <client-secret> --provider-scope-string "User.Read" --parameters tenantId=<tenant-id> tokenExchangeUrl=api://botid-<teamsAppId>
   az bot authsetting create -n <teamsAppId> -g <rg> -c graphmail --service Aadv2 --client-id <teamsAppId> --client-secret <client-secret> --provider-scope-string "Mail.Read" --parameters tenantId=<tenant-id> tokenExchangeUrl=api://botid-<teamsAppId>
   ```

6. Download the manifest, add:

   ```json
   "webApplicationInfo": {
     "id": "<teamsAppId>",
     "resource": "api://botid-<teamsAppId>"
   }
   ```

   then upload it again:

   ```powershell
   teams app manifest download <teamsAppId> manifest.json
   teams app manifest upload manifest.json <teamsAppId>
   teams app doctor <teamsAppId>
   ```

### Running the sample

```powershell
cd dotnet
dotnet build TeamsMiddleware.slnx
$env:ASPNETCORE_ENVIRONMENT = "Development"
dotnet run --project test_samples\TeamsMiddlewareSample --urls http://localhost:3980
```

### Direct Line / Web Chat

The shared harness lives in `tools\webchat\`.

```powershell
az bot directline show --name <teamsAppId> --resource-group <rg> --with-secrets -o json
py tools\webchat\serve.py
py tools\webchat\dl_test.py help channel "agents sdk react"
```

### Email

Email is intentionally **not** auto-configured by this repo because Azure requires a real mailbox
address and password:

```powershell
az bot email create -n <teamsAppId> -g <rg> -a <mailbox> -p <password>
```

Once enabled, expect auth commands on email to be declined with the channel-specific explanation.

## Testing matrix

| Channel | Input | Expected owner | Expected outcome |
|---|---|---|---|
| Teams chat | `help` | Teams SDK | Adaptive Card listing Teams commands |
| Teams chat | `react` | Teams SDK | Bot posts a message, adds 👍, then removes it |
| Teams chat | `quote` | Teams SDK | Quoted reply |
| Teams chat | `targeted` | Teams SDK | Targeted/private message |
| Teams chat | `task` | Teams SDK | Task module button, fetch, submit |
| Teams chat | react to bot message | Teams SDK | Reaction event summary |
| Teams chat | `channel` | Agents SDK | Reports `msteams` / subchannel |
| Teams chat | `whoami` | Agents SDK + OAuth | Sign-in confirmation, then Graph-backed user identity |
| Teams chat | `mail` | Agents SDK + OAuth | Sign-in confirmation, then Graph mail summary |
| Teams chat | `signout` | Agents SDK + OAuth | Sign-out confirmation |
| Teams chat | any other text | Agents SDK | `[Agent SDK]` echo |
| Web Chat / Direct Line | `help` | Agents SDK | Text help for non-Teams commands |
| Web Chat / Direct Line | `agents sdk react` | Agents SDK via Teams API client | Explains that the reactions API is unavailable on Direct Line |
| Web Chat / Direct Line | `agents sdk proactive` | Agents SDK via Teams API client | Sends a second message created through the Teams API client |
| Web Chat / Direct Line | `mail` / `whoami` | Agents SDK + OAuth | Auth prompt or sign-in confirmation plus Graph result, depending on channel support |
| Email | `help` | Agents SDK | First-line help text |
| Email | `whoami` / `mail` / `signout` | Agents SDK | Auth declined with email-specific explanation |

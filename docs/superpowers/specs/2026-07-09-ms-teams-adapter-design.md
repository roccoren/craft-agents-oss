# Microsoft Teams Messaging Channel — Design

**Date:** 2026-07-09
**Status:** Approved (design)
**Related:** Messaging Platform Expansion (follows the Discord adapter precedent,
`docs/superpowers/specs/2026-07-09-discord-adapter-design.md`)

## Summary

Add Microsoft Teams as a first-class messaging channel alongside Telegram,
WhatsApp, Lark, and Discord. A user connects an **Azure Bot** (Microsoft App ID +
App Password, optional Tenant ID); sessions can then be bound to Teams **1:1
personal chats**, **team channels**, and **group chats**. The adapter supports
text, message editing, inline buttons (Adaptive Card `Action.Submit`), and typing
indicators — matching the Telegram/Discord-tier experience (in-chat approvals via
buttons). File uploads are a documented v1 limitation.

## The core architectural difference

Every existing channel connects **outbound** (Telegram long-poll, WhatsApp/Discord
gateway WebSocket, Lark long-connection) and therefore advertises
`webhookSupport: false` and runs happily inside the Electron desktop app with no
public endpoint.

Teams is the opposite. The Bot Framework model is:

- **Inbound:** the Azure Bot Connector POSTs `Activity` JSON to a public HTTPS
  **messaging endpoint** (`/api/messages`), authenticated with a Bot Connector
  JWT.
- **Outbound:** the bot replies by calling the Bot Connector REST API using its
  app credentials — this half is a plain outbound HTTPS call and works from
  anywhere.

Consequences that drive the whole design:

1. **No subprocess worker.** Unlike Discord/WhatsApp (isolated because of native
   deps and persistent sockets), the `botbuilder` SDK is pure JS with no native
   modules and no persistent connection. The adapter runs **in-process**.
2. **A public URL is required.** We provide it with a bundled **Azure Dev
   Tunnels** persistent tunnel (stable URL) by default, or a user-supplied
   **bring-your-own** stable HTTPS URL.
3. **Replies are proactive.** The Router replies asynchronously, outside the
   inbound HTTP turn, so the adapter stores a `ConversationReference` per
   conversation and sends later via `continueConversationAsync`.

## Decisions (confirmed with user)

1. **Integration:** Bot Framework / Azure Bot Service (bidirectional bot).
2. **Runtime:** In-process adapter; **no** worker subprocess.
3. **Deployment:** Desktop-first with a bundled tunnel, and headless-capable.
4. **Tunnel:** Bundle **Azure Dev Tunnels** (`devtunnel`) with a **persistent
   (stable-URL) tunnel** as the default; also allow **BYO URL** override.
5. **Auth:** Microsoft App ID + App Password, **optional** Tenant ID. Tenant
   present ⇒ single-tenant app; absent ⇒ multi-tenant. Stored in the
   `messaging_bearer` credential row (`name = 'teams'`).
6. **Surfaces:** 1:1 personal chat, team channels, and group chats.
7. **Channel/group trigger:** Configurable per binding — `'mention'` (require
   @bot, default) or `'all'`. Personal chats always route.
8. **Approvals:** In-chat via Adaptive Card buttons (`approvalChannel: 'chat'`).

## Architecture

Two runtime halves, both in-process in the gateway host (Electron main or the
headless server):

```
Teams / Azure Bot Connector
        |  HTTPS POST  Activity  (Bot Connector JWT)
   Public tunnel URL  (devtunnel persistent  |  BYO)
        |
   Local HTTP listener (adapter-owned, http.createServer on 127.0.0.1:<port>)
        |
   TeamsAdapter  --- CloudAdapter.process --->  TurnContext
        |  translate + store ConversationReference
        |  implements PlatformAdapter
   Gateway -> Router -> SessionManager
        |
   reply (async):  continueConversationAsync(appId, ref, ctx => ctx.sendActivity)
        |  outbound HTTPS -> Bot Connector REST
   Teams
```

Router, `Commands` (`/new`, `/bind`, `/pair`), access-control, pending-senders,
and renderer chunking are platform-agnostic and reused unchanged.

### New adapter: `packages/messaging-gateway/src/adapters/teams/`

- `index.ts` — `TeamsAdapter implements PlatformAdapter`. Owns:
  - a `ConfigurationBotFrameworkAuthentication` + `CloudAdapter` from
    `botbuilder`, configured from the parsed credentials;
  - an adapter-owned local HTTP listener (Node `http.createServer`) that pipes
    requests into `CloudAdapter.process(req, res, logic)`;
  - a `Map<channelId, Partial<ConversationReference>>` for proactive replies,
    populated from every inbound turn (`TurnContext.getConversationReference`);
  - translation of activities to `IncomingMessage` / `ButtonPress`;
  - lifecycle: `initialize` (build adapter, start listener), `destroy` (close
    listener), `isConnected` (credentials validated + listener up).
- `format.ts` — Markdown → Teams: plain messages use Teams-supported markdown;
  buttons render as an Adaptive Card (`Action.Submit`, `data.buttonId`).

**Tenancy ("either"):** if `tenantId` is set, configure
`MicrosoftAppType = 'SingleTenant'` + `MicrosoftAppTenantId`; otherwise
`MicrosoftAppType = 'MultiTenant'`. The same field selection governs the token
authority used by `testTeamsCredentials`.

### Tunnel management: `packages/messaging-gateway/src/adapters/teams/tunnel/`

An isolated unit with a narrow interface so it can be tested and swapped:

```ts
interface TunnelProvider {
  start(localPort: number): Promise<{ publicUrl: string }>
  stop(): Promise<void>
  onUrlChange(cb: (publicUrl: string) => void): void
  isRunning(): boolean
}
```

- `DevTunnelProvider` — spawns `devtunnel host <tunnelId> --allow-anonymous`
  (persistent tunnel ⇒ **stable URL**), parses the public URL from stdout, keeps
  the process alive, restarts on unexpected exit. Requires a one-time
  `devtunnel user login` (Microsoft/GitHub); the binary is bundled via
  `extraResources`. The persistent tunnel id is stored in the workspace config.
- `ByoTunnelProvider` — no child process; returns the user-supplied stable HTTPS
  base URL as `publicUrl`.

The messaging endpoint shown to the user is `<publicUrl>/api/messages`; they
paste it into the Azure Bot resource's "Messaging endpoint" once. With a
persistent devtunnel or a BYO URL this value is stable across restarts.

### Credentials

- Stored in `messaging_bearer`, `name = 'teams'`, JSON
  `{ appId, appPassword, tenantId? }`.
- `parseTeamsCredentials(raw)` validates shape and returns
  `{ appId, appPassword, tenantId? }`.
- `testTeamsCredentials({ appId, appPassword, tenantId? })` performs an OAuth2
  client-credentials token request against the correct authority
  (`login.microsoftonline.com/<tenantId>` when single-tenant, else
  `login.microsoftonline.com/botframework.com`) with scope
  `https://api.botframework.com/.default`. A 200 with an `access_token`
  confirms the credentials without needing any inbound traffic.

## Incoming / outgoing mapping

### Inbound `Activity` → `IncomingMessage`

| IncomingMessage field | Source |
|-----------------------|--------|
| `platform`            | `'teams'` |
| `channelId`           | `activity.conversation.id` |
| `messageId`           | `activity.id` |
| `senderId`            | `activity.from.id` (AAD/Teams user id) |
| `senderName`          | `activity.from.name` |
| `text`                | `activity.text` with `<at>…</at>` mentions + leading bot mention stripped |
| `isDM`                | `activity.conversation.conversationType === 'personal'` |
| `mentionedBot`        | `true` if `activity.entities` contains a `mention` of the bot id |
| `attachments`         | v1: ignored (file upload/consent deferred) |
| `timestamp`           | `activity.timestamp` |
| `raw`                 | the `Activity` |

Every inbound turn stores
`conversationRefs.set(channelId, TurnContext.getConversationReference(activity))`.

### Button press

Adaptive Card `Action.Submit` arrives as a `message` activity with
`activity.value = { buttonId, data? }`. Mapped to `ButtonPress`
(`channelId`, `messageId`, `senderId`, `senderName`, `buttonId`, `data`). The
adapter acknowledges the turn (HTTP 200) so Teams does not show a failure.

### Outbound

All sends resolve the stored `ConversationReference` for `channelId` and run
inside `adapter.continueConversationAsync(appId, ref, async ctx => …)`:

- `sendText` → `ctx.sendActivity(text)`; returns the resulting activity id.
- `editMessage` → `ctx.updateActivity({ id, type: 'message', text })`.
- `sendButtons` → `ctx.sendActivity({ attachments: [adaptiveCard] })`.
- `sendTyping` → `ctx.sendActivity({ type: 'typing' })`.
- `clearButtons` → `ctx.updateActivity` replacing the card with plain text.
- `sendFile` → **v1 limitation:** posts the caption (and any URL) as text; true
  file upload (file-consent card / Graph) is deferred.

## Type-system & registry wiring

### `messaging-gateway/src/types.ts`

- `PlatformType` += `'teams'`.
- `AdapterCapabilities.markdown` union += `'teams'`.
- `MessagingConfig.platforms` += `teams?: { enabled: boolean; tunnelMode:
  'devtunnel' | 'byo'; byoUrl?: string; devtunnelId?: string;
  messagingEndpoint?: string }`.
- `BindingConfig` += `teamsChannelTrigger: 'mention' | 'all'` (default
  `'mention'`), handled in `normalizeBindingConfig` with a migration default;
  mirrors `discordGuildTrigger`. Personal chats always route regardless.
- `getDefaultBindingConfig`: Teams supports buttons ⇒ `approvalChannel` default
  `'chat'` (like Telegram/Discord, unlike WhatsApp's `'app'`).
- Generalise the `mentionedBot` / `isDM` doc comments (currently "Discord only")
  to "Discord and Teams".

### `registry.ts`

- `WorkspaceState` += `teams: TeamsAdapter | null` (+ any tunnel handle needed).
- `MessagingGatewayRegistryOptions` += `teams?: { devtunnelBin?: string;
  localPort?: number }`.
- New: `testTeamsCredentials`, `saveTeamsCredentials`, `connectTeams` /
  `tryConnectTeams`, `parseTeamsCredentials`.
- Add `'teams'` to every platform-iteration site: `initializeWorkspace`,
  `disconnectPlatform`, `forget`, runtime-info clone, and the
  `['telegram','whatsapp','lark','discord']` arrays.
- On connect: build the adapter, start its local listener, start the configured
  `TunnelProvider`, persist the computed `messagingEndpoint` into config, and
  surface it back to the caller (for the UI to display).

### `index.ts` / `package.json`

- Export `TeamsAdapter` + public types.
- Add `botbuilder` (and its transitive `botframework-connector`) as a dependency
  of `messaging-gateway`. No separate worker package.

### Electron `main/index.ts` + `electron-builder.yml`

- Provide `teams.devtunnelBin` (dev: repo-local vendored binary or PATH lookup;
  packaged: `resourcesPath/devtunnel/devtunnel[.exe]`) and a local webhook port.
- Add the `devtunnel` binary to `extraResources` (per-platform).
- Start/stop the tunnel with the adapter lifecycle.

### Trigger filter

`TeamsAdapter.onMessage` handling of an inbound activity:

- `personal` conversation → forward.
- `channel` / `groupChat` → look up the bound `teamsChannelTrigger`:
  `'mention'` forwards only when `mentionedBot === true`; `'all'` forwards every
  message. Pre-bind command messages require an @mention.

## Shared protocol + server-core + UI

### `shared/src/protocol/channels.ts`

- `RPC_CHANNELS.messaging` += `TEST_TEAMS: 'messaging:testTeams'`,
  `SAVE_TEAMS: 'messaging:saveTeams'`.
- `SAVE_TEAMS` returns the connection result including the computed
  `messagingEndpoint` so the dialog can display the URL to paste into Azure.
- Classify both channels in `routing.ts` (exhaustive routing test).

### `server-core`

- `messaging-registry-interface.ts`: add
  `testTeamsCredentials(creds: { appId; appPassword; tenantId? })` and
  `saveTeamsCredentials(workspaceId, creds, options)` returning the endpoint.
- `handlers/rpc/messaging.ts`: register `TEST_TEAMS` / `SAVE_TEAMS` handlers
  (mirrors the two Lark handlers).

### Electron renderer

- `transport/channel-map.ts`: add `testTeamsCredentials` /
  `saveTeamsCredentials`.
- New `components/messaging/TeamsConnectDialog.tsx`: fields for **App ID**,
  **App Password**, **Tenant ID (optional)**, a **tunnel mode** toggle
  (Dev Tunnel | Bring-your-own URL) with a conditional **BYO URL** field, plus
  help text ("create an Azure Bot", "run `devtunnel user login`", "paste this
  messaging endpoint into Azure"). Shows the computed messaging endpoint after
  connect. Test / Save actions.
- New `assets/messaging-icons/teams.svg`; register in
  `MessagingPlatformIcon.tsx` (bg `#6264A7`, initial `T`).
- Update the hardcoded platform unions that currently read
  `'telegram' | 'whatsapp' | 'lark' | 'discord'`:
  `MessagingSettingsPage.tsx`, `MessagingPlatformIcon.tsx`,
  `MessagingSessionMenuItem.tsx`, `PairingCodeDialog.tsx`, `atoms/messaging.ts`,
  `playground/mock-utils.ts`.
- `MessagingSettingsPage.tsx`: add a Teams platform card.

### i18n

- Add the Teams keys to `en.json` and every other locale file (parity enforced).
  Keep "Microsoft Teams" / "Teams" in English.

### Docs

- `doc-links.ts` + `word-lists.ts`: add the Teams doc link + session word-list
  entry (following the Lark/Discord precedent).

## Error handling

- Invalid credentials → `testTeamsCredentials` token request fails → dialog error.
- Missing / mis-set Azure messaging endpoint → inbound never arrives; the UI
  shows the exact endpoint to paste and a "waiting for first message" hint.
- `devtunnel` not logged in / not installed → `TunnelProvider.start` rejects →
  runtime `error` + dialog prompt to run `devtunnel user login`.
- Tunnel process exit → auto-restart with backoff; persistent tunnel keeps the
  same URL, so no Azure re-config needed.
- Proactive send with no stored `ConversationReference` (bot never saw the
  conversation) → surfaced as a send error.

## Testing

- adapter: `lifecycle.test.ts` — build adapter with mocked `CloudAdapter`,
  feed a synthetic `Activity`, assert `IncomingMessage` mapping + stored
  conversation reference; assert proactive `sendText`/`editMessage` call the
  mocked context; trigger-filter tests (mention / all / personal).
- format: `format.test.ts` — Markdown → Teams text + Adaptive Card button card.
- tunnel: `tunnel.test.ts` — `ByoTunnelProvider` returns the URL as-is;
  `DevTunnelProvider` parses the public URL from mocked spawn stdout and
  restarts on exit.
- credentials: `testTeamsCredentials` against a mocked token endpoint
  (single-tenant vs multi-tenant authority selection).
- No network, no real Azure Bot, no real tunnel.

## Verification

- Per package: `bun run tsc --noEmit` / `typecheck`.
- Root: `bun run validate:ci` (includes i18n parity / sorted / coverage).

## Phasing (single spec, phased plan)

1. **Core adapter + BYO URL + wiring + UI** — personal chat, text/edit/typing/
   Adaptive-Card buttons, credentials, registry + protocol + server-core + UI +
   i18n + docs. Works headless and desktop when the user supplies a stable URL.
2. **Bundled `devtunnel` persistent tunnel** — `DevTunnelProvider`, binary
   bundling, lifecycle wiring, dialog tunnel-mode UI.
3. **Channel + group surfaces** — `teamsChannelTrigger`, mention detection,
   binding UI copy.

Phases 1–3 constitute v1.

## Out of scope (v1)

- File uploads / file-consent card / Graph attachment flow.
- Slash commands, message extensions, task modules, SSO, tabs.
- Reactions, meeting/call events, presence.
- Auto-provisioning the Azure Bot resource or auto-updating its messaging
  endpoint via ARM.

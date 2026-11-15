# DSH Harness API recon — for a BrandDeck (PPTX deck generator) plugin bundle

Read-only recon over the extracted installation at
`D:\Projects\ai\BrandDeck\.dsh-src\dsh\node_modules\@deepseek-ai\<package>\lib\`.
Every claim below is quoted from that artifact (version `0.2.1-rc.3`). Where the
artifacts do not decide a question I write **unknown** and list what I ruled out.

Note on package availability: `@deepseek-ai/dsh-client-ui-slots` is **not installed**
(it is a `devDependencies` entry only — `dsh-client-ui-schedule/package.json` lists it
under `devDependencies`, and the directory does not exist on disk). The slot contract
therefore has to be read from (a) the generated client catalog inside
`dsh-cordis-client-runner/lib/client.js` (which carries the original `source:` paths and
the register-option docs) and (b) shipped consumers. That is what I quote.

---

## 1. CLIENT → HOST TRANSPORT

### 1a. `host.call` / `harness.handle` are dynamic-package-only. Not available to bundle client halves.

The dynamic runner evaluates a package's browser half as a **function body** and passes
`host` and `harness` as closure parameters:

`dsh-cordis-client-runner/lib/client.js`
```js
async function evaluateClientHalf(pluginId, clientCode, env, styles) {
	const traps = closureTraps();
	const parameters = [
		"React",
		"console",
		"styles",
		"host",
		"harness",
		...Object.keys(traps),
		"process",
		"Buffer"
```

The dynamic half is explicitly a *sealed* closure — `fetch` is a throwing trap that
redirects to `host.call`:

```js
const DYNAMIC_CLIENT_REDIRECTS = {
	setTimeout: TIMER_REDIRECT,
	setInterval: TIMER_REDIRECT,
	clearTimeout: TIMER_REDIRECT,
	clearInterval: TIMER_REDIRECT,
	fetch: "network belongs to the HOST half: register a handler there with harness.handle(method, fn) and call it here via host.call(method, args).",
	require: "modules cannot be imported here. React arrives as the `React` closure symbol; everything else goes through ctx services or host.call."
};
```
```js
function harnessTrap() {
	return new Proxy({}, { get(_target, prop) {
		throw new Error(`harness.${String(prop)} belongs to the HOST half (\`code\`): register handlers there with harness.handle(method, fn); the browser half calls them via host.call(method, args).`);
	} });
}
```

A **bundle** client half is different. It is a plain CJS factory executed by the page's
module system, and the only thing it receives is `require`:

`dsh-client-modules/lib/index.js`
```js
 * Lazy CJS model: executing a plugin bundle only REGISTERS its
 * factory (`window.__ModuleLoader__.load({id, factory})`); every module body
 * side effect — including CSS injection — lives inside the factory closure
 * and runs at materialization, not at script execution.
```

There is no `host`/`harness` entry in the builtin `require` table either — the table is
fixed (`dsh-web-frontend/dist/assets/index-BXV-oxvX.js`):
```js
{react:aC,"react/jsx-runtime":dC,"react-dom":gC,"react-dom/client":_C,"@deepseek-ai/cordis":Ub,zod:Vj,
 "@tanstack/react-virtual":Xj,"@deepseek-ai/dsh-client-store":xI,"@deepseek-ai/dsh-client-ui-slots":jI,
 "@deepseek-ai/dsh-client-ui-primitives":TU,"@deepseek-ai/dsh-client-ui-dockkit":SZ}
```

**Verdict:** `host.call`/`harness.handle` is exclusive to `cordis_run`-style dynamic packages.
A bundle client half has **exactly two** channels to the host: (1) whatever Cordis
services the page composition provides (`ctx.remote.*`, `ctx.connection`, …), and (2) plain
browser `fetch` against HTTP routes the host half registers.

### 1b. HTTP routes: raw registration has no auth; `ctx.connection` owns the fence

`dsh-host-webserver/lib/index.js` — `WebRoute` is `{kind, path, handler}`, mounted at
whatever absolute path you give it (no prefix, no auth):

```js
	register(route) {
		const table = route.kind === "exact" ? this.exact : this.prefixes;
		if (table.has(route.path)) throw new Error(`webserver: duplicate ${route.kind} route "${route.path}"`);
		table.set(route.path, route);
		return () => {
			table.delete(route.path);
		};
	}
```
```js
			const rawPath = new URL(req.url ?? "/", "http://x").pathname;
			const route = this.match(rawPath);
			if (route !== void 0) {
				await route.handler(req, res);
				return;
			}
```
Handler signature is raw `node:http` — the route owns the response to completion:
`handler(req, res)`. Prefix matching is longest-prefix-wins (`match()`), `registerUpgrade`
exists for WebSocket upgrades, `registerFallback` is single-owner.

There is **no** `admit` / `authorizeIndex` / `requestRejection` anywhere in the webserver.
Those live on the `connection` service (`super(ctx, "connection")`), and the shipped
`/api` prefix route applies them:

`dsh-client-connection/lib/index.js`
```js
	requestRejection(request) {
		const admission = this.admit(request);
		return "rejection" in admission ? admission.rejection : void 0;
	}
	/** Browser requests share an owner Peer; API HTTP does not depend on an attached socket. */
	admit(request) {
		if (!this.isTrusted(request)) return { rejection: 403 };
		if (this.disposed) return { rejection: 401 };
		const secure = this.owner.get("webServer")?.protocol === "https:";
		if (this.browserAuth.isAuthenticated(request, secure)) return { peer: this.browserPeer };
		const id = this.browserAuth.authenticateApi(request, secure);
		const record = id === void 0 ? void 0 : this.peers.get(id);
		return record === void 0 ? { rejection: 401 } : { peer: record.peer };
	}
```

The `Host`/`Origin` fence is `isTrustedApiRequest` — this is the function that decides
whether a same-origin page fetch is allowed:

```js
function isTrustedApiRequest(request, trustedHosts, bindHost, protocol = "http:") {
	const host = header$1(request.headers, "host");
	if (host === void 0) return false;
	const hostUrl = parseAuthority(host, protocol);
	if (hostUrl === void 0) return false;
	if (!isLoopbackHostname(hostUrl.hostname) && !isTrustedAuthority(hostUrl, trustedHosts, protocol) && !isBindAddressAuthority(hostUrl, bindHost)) return false;
	if (header$1(request.headers, "sec-fetch-site") === "cross-site") return false;
	const origin = header$1(request.headers, "origin");
	if (origin === void 0) return true;
	return isSameAuthorityOrigin(origin, host);
}
```

and the cookie it checks is minted by `authorizeIndex` on the launch-token exchange and is
`HttpOnly; SameSite=Strict; Path=/`:

```js
function sessionCookie(name, value, expiresAt, maxAgeSeconds, secure) {
	const attributes = `${name}=${value}; Max-Age=${String(maxAgeSeconds)}; Path=/; Expires=${new Date(expiresAt).toUTCString()}; HttpOnly; SameSite=Strict`;
	return secure ? `${attributes}; Secure` : attributes;
}
```
```js
const cookieName = (authority) => COOKIE_PREFIX + encodeBase64Url(createHash("sha256").update(authority).digest());
// COOKIE_PREFIX = "dsh-auth-"
```

**Is a same-origin `fetch('/branddeck/api', {credentials:'include'})` from the page plausibly
authorized?** Yes, if the route itself performs the check. For the GUI at
`http://127.0.0.1:5664` the request carries `Host: 127.0.0.1:5664` (loopback ⇒ Host fence
passes), `Sec-Fetch-Site: same-origin` (≠ `cross-site` ⇒ passes), `Origin:
http://127.0.0.1:5664` (equal authority ⇒ passes), and the `dsh-auth-<sha256(authority)>`
cookie is sent automatically on a same-origin request (`SameSite=Strict` permits it;
`HttpOnly` only hides it from JS). `admit()` then returns `{peer: browserPeer}` whose
principal is `{kind:"owner"}`.

But note: the webserver itself does **not** call `admit` for plugin routes. A plugin route
must either do it itself —

`dsh-experimental-inspector/lib/index.js` (the shipped precedent for a plugin-owned route):
```js
					if (!ctx.connection.authorizeIndex(req, res)) return;
```
```js
					const rejection = ctx.connection.requestRejection(req);
					if (rejection !== void 0) {
						const reason = rejection === 401 ? "Unauthorized" : "Forbidden";
						socket.end(`HTTP/1.1 ${rejection} ${reason}\r\nConnection: close\r\nContent-Length: 0\r\n\r\n`);
						return;
					}
```

— or mount under `/api`, where the connection layer already admitted the request:

```js
		ctx.disposers.push(ctx.connection.fetch.register({
			path: INSPECTOR_BOOTSTRAP_PATH,          // "/api/experimental-inspector/bootstrap"
			methods: ["GET"],
			requestBody: "buffered",
			fetch: (request) => bootstrapResponse(ready.promise, request.signal)
		}));
```

`fetch.register` is an exact-route registry whose paths are constrained to `/api/…`, and
dispatch additionally requires an owner principal:

```js
	assertFetchRoute(route) {
		if (endpointFromPath("/api", route.path) === void 0) throw new Error(`connection: invalid exact Fetch route ${JSON.stringify(route.path)}`);
```
```js
			fetch: (request, peer) => {
				const pathname = new URL(request.url).pathname;
				const route = this.fetchRoutes.get(pathname);
				if (route?.methods.has(request.method) === true) {
					if (peer.principal.kind !== "owner") return Promise.resolve(new Response("forbidden", { status: 403 }));
					return route.fetch(request);
				}
```

### 1c. `ctx.remote` exists, but plugin-authored namespaces need build-time codegen

`ctx.remote` is provided by the **api-gateway client half**
(`dsh-api-gateway/lib/client.browser.js`):
```js
		const inject = ["typert", "connection"];
		function apply(ctx) {
			new ClientRemoteService(ctx);
		}
		var ClientRemoteService = class extends _deepseek_ai_cordis.Service {
			constructor(ctx) {
				super(ctx, "remote");
				this.ownerCtx = ctx;
				const connection = ctx.get("connection");
```
Its own surface is only `$stream`, `$host`, `$mount`, `$on`:
```js
			async $mount(contribution) {
				const callerCtx = this.ctx;
				const owned = callerCtx.effect(async () => {
					const dispose = await this.enqueue(() => this.mountContribution(callerCtx, contribution));
```
`$mount` consumes a **generated Typert contribution** (`createTypertRemoteInfo({package, methods:{…}})`
with per-method `parameters`, `codec`, `jsonSchema`), which is compiler output shipped as
`lib/typert.remote-client.js`. A plain-JS bundle cannot author one by hand at reasonable cost.

A plain-JS client half **can** read the service (`inject: ['remote']` or `ctx.get('remote')`),
and one namespace it exposes is `commands` — mounted by `@deepseek-ai/dsh-api-remotes`
(`disposers.push(await ctx.remote.$mount(contribution))`, based on the 1 MB generated
`lib/client.js`). The generated descriptor for it is:

`dsh-api-remotes/lib/client.js`
```js
		package: "@deepseek-ai/dsh-commands",
		methods: {
			"commands/execute": {
				service: "commands",
				invocation: { kind: "direct" },
				scope: {
					context: "agent",
					wire: "agentId"
				},
				parameters: [
					{ name: "agent", wire: "agentId", description: "exact receiving agent.", source: "lookup", lookup: "agent", … },
					{ name: "line", description: "complete slash-command line.", source: "json", … },
					{ name: "submittedAttachments", description: "encoded images and staged file receipts…", source: "json", … },
				]
```
and the host-side contract marks it Remote:

`dsh-tool-cordis/lib/types/api-catalog.js` (`key: 'commands'`)
```js
                signature: '@Remote async execute( agent: Agent, line: string, submittedAttachments: readonly CommandSubmitAttachment[], signal: AbortSignal, ): Promise<CommandExecution | undefined>',
```
Shipped client call sites show the concrete argument shape — the `agent` param is passed as a
**session id string**, attachments as `[]`:

`dsh-client-ui-commands/lib/client.js`
```js
				const result = await this.ctx.remote.commands.execute(session.sessionId, line, attachments);
```
`dsh-client-ui-plan/lib/client.js`
```js
					const result = await ctx.remote.commands.execute(sessionId, "/plan off", []);
```

**So:** `ctx.remote.commands.execute(sessionId, '/branddeck export', [])` is reachable from a
bundle client half **without codegen**, but it only invokes *registered slash commands*. It is
not a general host RPC. It is also the pattern the first-party plugin skill recommends — see
`dsh-agent-preset/skills/cordis-plugin-development/references/user-actions.md`:

> 2. The UI action calls it through a Client-callable Host entry point found with inspection,
>    such as a session command that `ctx.remote.commands.execute()` runs, and shows the returned failure.

### 1d. RECOMMENDED transport: the generic Connection RPC channel

Undocumented in the skill references but first-class and symmetric in the shipped code: the
`connection` service carries a **context-scoped generic unary RPC registry** on both halves.

Host (`dsh-client-connection/lib/index.js`):
```js
	/** Generic channel registry scoped to the Context reading this service. */
	get rpc() {
		const owner = this.ctx;
		return {
			handle: (channel, handler) => this.register(owner, channel, handler),
			intercept: (channel, matches, handler) => this.registerInterceptor(owner, channel, matches, handler)
		};
	}
```
```js
	register(owner, channel, handler) {
		assertChannel(channel);
		const fetchHandler = rpcFetchHandler(channel, handler);
		const route = {
			kind: "prefix",
			path: channel,
			handler: async (req, res) => {
				const admission = this.admit(req);
				if ("rejection" in admission) {
					res.writeHead(admission.rejection);
					res.end(admission.rejection === 401 ? "unauthorized" : "forbidden");
					return;
				}
				await bridge(req, res, fetchHandler, admission.peer);
			}
		};
		return owner.effect(() => owner.webServer.register(route), `client-connection: ${channel} rpc channel`);
	}
```
The handler contract and the envelope validation:
```js
	async fetch(request, peer) {
		const endpoint = endpointFromPath(channel, new URL(request.url).pathname);
		if (request.method !== "POST" || endpoint === void 0) return new Response("not found", { status: 404 });
		if (request.headers.get("content-type")?.split(";", 1)[0]?.trim().toLowerCase() !== "application/json") return new Response("content type must be application/json", { status: 415 });
		…
			const result = await handler(endpoint, message.payload, request.signal, peer);
			return fullResponse(message.rpcId, result);
```
Channel/path grammar (note `/api` is reserved for the shared channel):
```js
const CHANNEL_PATTERN = /^\/[A-Za-z0-9._~-]+$/;
const ENDPOINT_SEGMENT_PATTERN = /^[A-Za-z0-9_$.-]+$/;
…
function assertChannel(channel) {
	if (!CHANNEL_PATTERN.test(channel) || channel === "/api") throw new Error(`connection: invalid or reserved RPC channel ${JSON.stringify(channel)}`);
}
```
Client side — the page **already** has this object as `ctx.connection`
(`dsh-client-connection/lib/client.browser.js`):
```js
		function createWebConnectionRpc(doFetch, openStream) {
			const send = doFetch ?? ((input, init) => globalThis.fetch(input, init));
			return {
				async call(channel, endpoint, payload, signal) {
					assertTarget(channel, endpoint);
					const rpcId = RpcId(randomUuid());
					const message = { type: "client-request", rpcId, method: endpoint, payload };
					const response = await send(`${channel}/${endpoint}`.slice(1), {
						method: "POST",
						headers: { "content-type": "application/json" },
						body: JSON.stringify(message),
						...signal === void 0 ? {} : { signal }
					});
					if (!response.ok) throw new Error(`transport failure for ${channel}/${endpoint}: HTTP ${response.status}`);
					…
					if (full.rpcId !== rpcId) throw new Error(`rpcId mismatch for ${endpoint}: sent ${rpcId}, got ${full.rpcId}`);
					return full.result;
				},
```
```js
			ctx.provide("connection", handle);
```
and `handle` includes `rpc`:
```js
			const handle = {
				interactions,
				openWebSocket: (path) => openConnectionWebSocket(path, …),
				isLoopback: …,
				generation: { getSnapshot: …, subscribe: … },
				state: { … },
				rpc,
```

**Why this is the recommended transport**
- It is a supported seam on both halves (`get rpc()` is documented as public, and the client
  half is a `ctx.provide`d service), used by the shipped gateway for `/api`.
- Authentication/authorization is **automatic**: `register()` calls `this.admit(req)` and
  answers 401/403 itself, so the plugin writes zero auth code and cannot get it wrong.
- No build tooling (no Typert codegen), pure JS, works from `window.__ModuleLoader__` factories.
- Registration is an `owner.effect(...)`, so it is disposed with the plugin fiber (HMR/unload safe).
- The handler receives the admitted `peer`, so a plugin can check `peer.principal.kind === 'owner'`.
- Request/response is JSON with a correlated envelope and an `ok|error` result union — no need
  to hand-roll status codes.
- Not from a session: unlike `commands.execute`, it does not require a live agent/session,
  which matters because our panel lives in the root-scoped `main` slot (see §2b).

Costs/caveats: it is not covered by the skill's `user-actions.md` guidance; the wire format is
an internal envelope (fine — both ends are the same code you ship); the client call resolves
`${channel}/${endpoint}` **relatively** (`slice(1)`), i.e. the page must sit at a directory
root, exactly as the shipped `/api` RPC already assumes.

---

## 2. SLOTS

### 2a. Exact registration calls (from shipped consumers)

`dsh-client-ui-schedule/lib/client.js` — the reference full-page plugin:
```js
			ctx.slots.inject("main", () => ctx.slots.register({
				name: "main",
				key: PANEL_ID,
				locale: MANAGER_NS,
				inject: () => ({
					...detail,
					onNewTask: () => {
						ctx.uiWorkspace.startSession();
					}
				})
			}, TaskManagerPage));
			ctx.slots.inject("sidebar.panellist", () => ctx.slots.register({
				name: "sidebar.panellist",
				id: PANEL_ID,
				order: 10,
				locale: MANAGER_NS,
				label: () => t("panel")
			}, TaskManagerIcon));
```
`dsh-client-ui-plugin-manager/lib/client.js`:
```js
			ctx.slots.inject("sidebar.panellist", () => ctx.slots.register({
				name: "sidebar.panellist",
				id: PANEL_ID,
				order: 0,
				label: () => t("panel"),
				locale: NS
			}, PluginsPanelIcon));
```

**Key point:** the keyed slot uses **`key`**; list slots use **`id`**. Passing the wrong one is
the most likely silent failure.

### 2b. `main` is keyed + root-scoped, and receives NO `sessionId`

From the generated client slot catalog (`dsh-cordis-client-runner/lib/client.js`):
```js
			{
				key: "main",
				kind: "keyed",
				scope: "root",
				summary: "Central panel selected by sidebar entry id.",
				doc: "Central panel selected by sidebar entry id. The reserved `conversation`\nkey hosts the Conversation; other keys receive no Session binding.",
				registerOptions: [{
					name: "key",
					requirement: "required",
					type: "string",
					doc: "Your cell key: the entry renders where the owner dispatches this exact key. Registering an already-occupied key replaces that occupant."
				}],
				ownerProps: [],
				standardProps: [
					"useRemoteAccess: (action: RemoteEndpoint) => boolean",
					"useResource: UseResource",
					"useWorkspaces: SnapshotSelectorHook<WorkspaceSnapshot>",
					"usePanelInfo: UsePanelInfo",
					"useSessions: UseSessions",
					"useSessionStatus: UseSessionStatus",
					"useSessionRetainInfo: UseSessionRetainInfo",
					"useSessionAccess: (sessionId: SessionId, action: RemoteEndpoint) => boolean",
					"useWorkspaceAccess: (workspaceId: WorkspaceId, action: RemoteEndpoint) => boolean"
				],
				keyDomain: "open: any string the owner dispatches (no compile-time key set), already taken: conversation",
				declaredBy: "an entry in 'root' (client-ui-layout), so it exists while that entry is mounted",
				occupants: [
					"client-ui-conversation ConversationPanel key 'conversation'",
					"client-ui-plugin-manager PluginManagerPage",
					"client-ui-schedule TaskManagerPage"
				],
				replaceRisk: "shadows-shipped-ui",
				source: "packages/client/ui-layout/src/client/index.ts:73"
			},
```
So a `main` component receives the nine **framework shares** (hooks) — `usePanelInfo`,
`useSessions`, `useWorkspaces`, … — plus whatever your `inject` factory returns. There is no
`session: SessionId` prop. (`conversation` is the exception: its `main.conversation` child slot
is `scope: "session-maybe"` and its `standardProps` do include `sessionId: SessionId | undefined`.)
Contrast `sidebar.panellist`, which is a list slot with `ownerProps`:
```
export interface SidebarPanelIconOwnerProps {
  size: number   // Requested square edge in pixels.
  active: boolean // Whether this panel is selected in the main column.
}
```

The owner of `main` is `@deepseek-ai/dsh-client-ui-layout`, which renders exactly one keyed entry:
```js
		/** Subscribe to the main key without subscribing the column frame to each panel id. */
		function MainPanel({ usePanelInfo, renderSlot }) {
			return renderSlot("main", {}, { entryKey: usePanelInfo((info) => info.activePanelId) ?? "conversation" });
		}
```
The panel controller lives in the same file:
```js
					instance.actions.retainMainPanels(ctx.slots.entries("main").flatMap((entry) => entry.options.key === void 0 ? [] : [entry.options.key]));
				…
				const layout = new LayoutController(instance.actions, (id) => ctx.slots.entries("main").some((entry) => entry.options.key === id), {
```
and rejects unregistered ids:
```js
				if (panelId !== null && !this.hasMainPanel(panelId)) throw new Error(`layout.selectPanel: main panel "${panelId}" is not registered`);
```

**Consequence for BrandDeck:** the deck panel cannot rely on a session prop. Either scope the
deck store by workspace (recommended — see §7), or read the active session through
`ctx.sessions`/`useSessions` hooks.

### 2c. `ctx.slots.inject(ownerKey, cb)` and the locale pattern

Semantics (client catalog, `key: "slots"`):
```js
						signature: "inject(key: keyof SlotMap & string, callback: () => SlotInjectionEffect): () => void",
						description: "Install an effect for each declaration lifetime of a slot. The callback runs synchronously when the declaration already exists; otherwise it runs inside the declaring `register()` call after the declaration is committed. Collapse disposes the effect and a later declaration runs it again. … The controller belongs to the caller's fiber, so plugin unload cancels a pending wait and removes any active contribution.",
```
Practices doc states the same idiom
(`dsh-agent-preset/skills/cordis-plugin-development/references/practices.md`):
> Contribute through slots: `ctx.slots.inject(ownerKey, () => ctx.slots.register(...))`. The callback's registrations are disposed when the owning declaration collapses and reinstalled when it returns.

Locale: **`register` the namespace, then `bind` it for `t`.** Shipped pattern
(`dsh-client-ui-schedule/lib/client.js`):
```js
		const MANAGER_NS = "schedule.manager";
		const PANEL_ID = "schedules";
		const inject = [ "slots", "locale", "remote", "remote.schedule", …, "sessions", "workspaces", "sidebarRightTabs", "sidebarRight" ];

		function apply(ctx) {
			ctx.effect(() => ctx.locale.register(NS, {
				zh: zh$1,
				en: en$1
			}), "ui-schedule: dictionaries");
			ctx.effect(() => ctx.locale.register(MANAGER_NS, {
				zh,
				en
			}), "ui-schedule: manager dictionaries");
			const t = ctx.locale.bind(MANAGER_NS);
```
Service contract (client catalog, `key: "locale"`):
```js
						signature: "register<N extends Extract<keyof LocaleNamespaceMap, string>>(ns: N, dicts: Record<BuiltInLocaleId, LocaleDictOf<N>>): () => void",
						description: "Register a declared namespace's dictionaries, all locales in one call — the typed form … every shipped locale is required (bilingual balance enforced at registration). Duplicate (ns, locale) throws …"
						signature: "register(ns: string, locale: string, dict: LocaleDict): () => void",
						description: "Single-locale untyped form for language-pack contributions and namespaces outside the merge table.",
```
`locale.bind(ns)` is the shipped way to obtain `t`. **unknown:** whether the slot option
`locale: NS` additionally injects a `t` prop into the registered component — I could not prove
it from the artifacts (`N extends keyof LocaleNamespaceMap` is a *type-level* constraint; the
runtime validation lives in the un-installed `ui-slots` package). Safe approach for a plain-JS
plugin: keep using your own `const t = ctx.locale.bind(NS)` and pass it through `inject: () => ({ t })`.

### 2d. `SlotCore['register']` — the full option surface

Verbatim declaration (client catalog, `key: "slots"` → `SlotCore`):
```js
export class SlotCore {
    readonly registerFactory: RegisterFactory;
    factory(name: string): StoredFactory | undefined;
    factoryVersion(name: string): number;
    subscribeFactory(name: string, listener: () => void): () => void;
    isFactoryLive(definition: StoredFactory): boolean;
    register<K extends keyof SlotMap & string, const EntryKey extends EntryKeyOf<K> = EntryKeyOf<K>,
             const D extends ChildrenDecl = Record<never, never>, H extends StoreDecl | undefined = undefined,
             M = never, N extends (keyof LocaleNamespaceMap & string) | undefined = undefined,
             C extends SlotComponent<never> = SlotComponent<never>>(
        options: BaseOptions<K, EntryKey, D, H, M, N> & { inject?: undefined },
        component: C & SlotComponent<ComposedProps<K, NoInfer<EntryKey>, keyof NoInfer<D> & keyof SlotMap & string, HandleOf<NoInfer<H>>, object, NoInfer<M>, NoInfer<N>>> & RendersCheck<C, D>
    ): () => void;
    register<…, I extends object, …>(
        options: BaseOptions<K, EntryKey, D, H, M, N> & { inject: (...args: InjectParams<K, H>) => I },
        component: C & SlotComponent<ComposedProps<…, I, …>> & RendersCheck<C, D>
    ): () => void;
    register(options: ErasedOptions, component: unknown): () => void;
    isLive(entry: StoredEntry): boolean;
    entries(key: string): readonly StoredEntry[];
    entriesOfSlot(key: string): readonly StoredEntry[];
    spec<K extends keyof SlotMap & string>(key: K): SlotSpec<SlotMap[K]> | undefined;
    specDynamic(key: string): SlotSpec<SlotEntryDef> | undefined;
```
Two overloads: the `inject` factory is optional, and when present its return value is added to
the component props (`I`). `BaseOptions` is **not expanded** anywhere in the artifacts
(`grep` for `interface BaseOptions` → no match; `ui-slots` is not installed), so its full field
list is **unknown**. Confirmed-used fields, from shipped registrations and the catalog:

| field | evidence | notes |
|---|---|---|
| `name` | every call | slot key; must match the `inject` key |
| `key` | `main` registerOptions | **required for keyed slots** |
| `id` | `sidebar.panellist` registerOptions | **required for list slots** |
| `order` | `sidebar.panellist` registerOptions | `"Position among the entries, ascending (default 0)."` |
| `label` | `sidebar.panellist` registerOptions | `"string | (() => string)"`, thunk re-read per projection |
| `locale` | ui-schedule / plugin-manager calls; `N extends keyof LocaleNamespaceMap` | namespace name |
| `inject` | ui-schedule (`inject: (sessionId) => ({…})` and `() => ({…})`) | props factory; some slots pass a context arg |
| `children` / store / `select` | only hinted by type params `D`, `H`, `M` | **unknown** — no shipped usage found; do not use |

### 2e. `ctx.sessions`, `ctx.theme`, `ctx.layout` — all present on the client ctx

Provided-service catalog (`CLIENT_SERVICE_API`, `dsh-cordis-client-runner/lib/client.js`):

`key: "layout"` — "Panel navigation and geometry actions exposed through ctx.layout.":
```js
						signature: "selectPanel(panelId: MainPanelId | null): void",
						description: "Select a global central panel without changing the current Session.",
						parameters: [{ name: "panelId", description: "registered main key, or null to show the Conversation." }],
						throws: ["if the selected main key is not registered; preserves the current selection."]
```
Shipped usage for exactly our case — sidebar icon → main panel
(`dsh-client-ui-sidebar/lib/client.js` line 598: `ctx.layout.selectPanel(id);`), and
`dsh-client-ui-plugin-manager/lib/client.js`:
```js
				const disposeNavigation = ctx.reflect.provide("pluginNavigation", {
					openList: () => {
						ctx.layout.selectPanel(PANEL_ID);
```
plus the active-panel read:
```js
				yield ctx.layout.panelInfo.subscribe(() => {
					if (ctx.layout.panelInfo.getSnapshot().activePanelId !== "plugins") instance.actions.setView({ kind: "list" });
				});
```

`key: "sessions"` — "The sessions-service face injected as `ctx.sessions`." (`retain`, `using`,
`retainInfo`, `refreshProjections`, `search`, `fork`, `scope`, `binding`).

`key: "theme"` — "Theme registry and preference owner." (`getTheme`, `setTheme`, `setFontSize`,
`register`, `overrideTokens`).

Injection is by service name in the plugin's `inject` array (ui-schedule: `["slots","locale",
"remote","sessions","workspaces","sidebarRightTabs","sidebarRight", …]`).

---

## 3. HOST TOOLS

### 3a. ToolDefinition

Registration is validated in `dsh-tools/lib/index.js`:
```js
	register(definition) {
		const name = definition.name;
		const output = definition.output;
		if (output === void 0 || typeof output !== "object" || typeof output.render !== "function" || output.presentationMeta !== void 0 && typeof output.presentationMeta !== "function") throw new TypeError(`tool "${name}" must declare output { schema, render, presentationMeta? }`);
		assertSupportedJsonSchema(output.schema);
		const timeoutMs = definition.timeoutMs;
		…
		if (name === "run_code") throw new Error(`tool name "${RUN_CODE_NAME}" is reserved for the PTC mode presentation transport and cannot be registered or shadowed`);
```
So `parameters` must be **JSON Schema** (not Zod, not schemastery). First-party packages author
a compact spec and compile it with `defineTool`:
```js
function defineTool(options) {
	…
	const parameters = parameterSchemaSpecToJsonSchema(options.parameters);
	const outputSchema = valueSchemaSpecToJsonSchema(options.output.schema);
	const validate = (args) => validateJsonSchemaValue(parameters, args, "");
```
A real definition (`dsh-tool-present/lib/index.js`) — this is the shape to copy:
```js
	ctx.tools.register(defineTool({
		name: "present",
		description: "Declare existing files as final deliverables for the user. …",
		parameters: { files: {
			type: "array",
			required: true,
			description: "Usually the 1-2 most important deliverables; at most 4 per call.",
			items: {
				type: "object",
				additionalProperties: false,
				properties: {
					path: { type: "string", required: true, description: "…" },
					description: { type: "string", description: "Brief description for the user." }
				}
			}
		} },
		output: {
			schema: { type: "object", additionalProperties: false, properties: { … } },
			render: (_args, value) => [{
				type: "text",
				text: value.files.map((file) => `Presented ${file.path}`).join("\n")
			}]
		},
		async execute(args, exec) {
			if (exec.agent === void 0) throw new Error("present requires an agent Session");
			…
			return { turn: boundary.lastTurn, files };
		}
	}));
```
`exec` carries the call context: `exec.agent`, `exec.signal`, `exec.callId` are used above.
Return value = the `output.schema`-shaped object; the model-visible text comes from
`output.render(args, value)`. Throwing produces an error result.

`dsh-tool-schedule` and `dsh-tool-present` both pass through `defineTool` from
`@deepseek-ai/dsh-tools`; shipping-as-import is acceptable because
`references/host-plugin.md` states: *"Packages shipped with dsh resolve from the dsh
installation, so the bundle declares no dependencies on them."*

**Do not pass the compact spec straight to `ctx.tools.register`.** `defineTool` is what
compiles it, and `register` stores the *compiled* result — proven by `schemaOf`, which refuses
anything that is not already lossless JSON Schema:
```js
	schemaOf(definition, detachParameters) {
		const { name, description, parameters, deferLoading } = definition;
		const detached = detachParameters ? snapshotJsonValue(parameters) : parameters;
		if (detached === void 0) throw new Error(`tool "${name}" parameters must be lossless JSON before schema projection`);
		return { name, description, parameters: detached, … };
```
So either wrap with `defineTool({...})`, or hand-write real JSON Schema:
```js
parameters: { type: "object", additionalProperties: false,
  properties: { action: { type: "string", enum: ["get", "put", "export"] } },
  required: ["action"] }
```
Note also that only `defineTool` adds argument validation (`validate` → `ToolArgsError`);
a hand-built definition gets no such gate.

### 3b. Availability, restriction, guards, inject

`dsh-tool-cordis/lib/types/api-catalog.js` (`key: 'tools'`):
```js
                signature: 'restrict(filter: ToolRestriction): () => void',
                description: 'Restrict global tools for the calling agent scope. Empty filters, unknown names, scope-local names, and reserved transport names fail. Restrictions intersect; scoped registrations remain visible.',
                signature: 'guard(guard: ToolGuard): () => void',
                description: 'Register a monotonic guard after the extensible `tools/pre-execute` waterfall. … Any matching guard may deny by returning a reason, while no guard can force-allow a call another guard denied.',
                signature: 'presentAs(mode: ToolPresentationMode): () => void',
```
"unavailable outside certain sessions" therefore = register on a narrower context, or
`ctx.tools.restrict(...)` / `guard(...)` from `agent.ctx`. `inject` is the ordinary Cordis
plugin `inject` export (`dsh-tool-present`):
```js
const inject = [ "tools", "fs", "sessionProjections", "workingDirectory" ];
```

### 3c. Yes — `ctx.tools.execute(...)` is a public method (correction to the brief's assumption)

```js
                signature: 'async execute(exec: ToolExecutionInput): Promise<ToolExecutionResult>',
                description: 'Execute through pre-policy, guards, around-dispatch, post-policy, definition-owned content finalization, and final notification. … an invisible tool reports `UNKNOWN_TOOL`.'
```
Implementation reads the input fields directly (`createExecution`):
```js
			const callId = exec.callId;
			const rootCallId = exec.rootCallId ?? callId;
			const name = exec.name;
			const agent = exec.agent;
			const parent = exec.parent;
			const signal = exec.signal;
			const visible = this.get(name, agent);
```
`agent` is optional (`...agent !== void 0 ? { agent } : {}`), so a host plugin (or a script with
a booted composition) can run its own tool in-process — e.g. reuse the exporter from the RPC
handler. It still needs a `callId` and a `signal`, and results are materialized snapshots
(`{ok:false, error:{code}}` for `UNKNOWN_TOOL`). A *bare* `node` script outside a booted Cordis
tree cannot call it at all (no `ctx`), so for testability prefer extracting the pure logic into
a plain function that both the tool and the RPC handler call.

---

## 4. SYSTEM PROMPT + LLM

### 4a. `ctx.systemPrompt.section(...)`

```js
	section(section) {
		if (!Number.isFinite(section.order)) throw new TypeError(`prompt section "${section.name}" order must be a finite number`);
		return this.layers.effect(this.ctx, (layer) => layer.sections.insert(section.name, section), { label: "systemPrompt.section()" });
	}
```
Section fields observed in the assembler: `name`, `text` (**string or `(context) => string`**),
`order` (**finite number, required**), `interpolate` (boolean; `false` keeps literal `{{…}}`),
`complete` (boolean; at most one active):
```js
					text: typeof section.text === "function" ? section.text(context) : section.text,
					...section.interpolate !== void 0 ? { interpolate: section.interpolate } : {}
```
Central placement constants for reference (`dsh-system-prompt/lib/index.js`):
```js
const SECTION_ORDERS = {
	HARNESS_IDENTITY: -1e3,
	DEPLOYMENT_PERSONA_PREFIX: 0,
	PLAN_POLICY: 500, TEAM_POLICY: 600, PTC_ONLY: 800, FILE_REFERENCE: 900,
	TOOL_BASH: 1e3, … TOOL_PRESENT-adjacent range …, TOOLS_SDK: 5e3,
	DELIVERABLE_FILE_REFERENCES: 9e3, STRUCTURED_OUTPUT: 9900,
	HARNESS_SOURCE: 1e4, WEB_SURFACE: 10100, DEPLOYMENT_PERSONA_SUFFIX: 10200
};
```
A plugin's own section should use an order between those bands (e.g. `2600`-ish or anything
unused) — the values are plain numbers, not an enum you must join.

### 4b. There is **no** `ctx.llm.complete(...)`. The LLM service is streaming-only.

`dsh-tool-cordis/lib/types/api-catalog.js` (`key: 'llm'`) — complete method list:
`registerAdapter`, `@Remote listProviders`, `registerConfigurableProviders`,
`@Remote listConfigurableProviders`, `registerModelDiscovery`, `@Remote discoverModels`,
`providerRetryPolicy`, `imageRequestPricing`, `imageRequestTarget`, `fileRequestText`,
`listModels`, `resolveModelInfo`, `resolveCallConfig`, `prepareCall`, and:
```js
                signature: 'stream(options: GenerateOptions): AsyncIterable<StreamChunk>',
                description: 'Stream one model call as raw chunks (token-level deltas). … Adapter selection, dispatch, and iteration failures become terminal `error` or `aborted` finish chunks; middleware, nested-call, cleanup, and consumer failures remain thrown.',
                parameters: [{ name: 'options', description: 'the full request; `options.provider` selects the adapter.' }],
```
`prepareCall` resolves the model and pins the adapter registration:
```js
                signature: 'async prepareCall( config: LlmCallConfig, signal?: AbortSignal, configure?: ConfigureCall, ): Promise<PreparedLlmCall>',
```
and returns `{ model, stream }`:
```js
	async prepareCall(provider, model, signal) {
		return {
			model: await this.resolveModel(provider, model, signal),
			stream: (options) => this.stream(options)
		};
	}
```
`GenerateOptions`/`LlmCallConfig` shape is **partially unknown** from the artifacts (it is a
type in `dsh-llm/lib/types/*`, erased at runtime); the implementation reads at least
`options.provider`, `options.model`, `options.messages`, `options.signal`, and optionally
`options.config`:
```js
				const adapterCall = await adapter.prepareCall(options.provider, options.model, options.signal);
```
```js
				let projectedMessages = resolvedOptions.messages;
```

**Which pipeline is actually possible?**
- **(i) Agent authors the deck JSON via tools — definitely possible** and the low-risk route:
  tools (§3) + a system-prompt section (§4a) + `ctx.fs` (§7). No LLM plumbing, no provider or
  model choice, no token accounting in your plugin, and it keeps the session log as the record.
- **(ii) The plugin calls the LLM itself — possible, but only via `ctx.llm.stream()`.**
  It is a supported service method (any plugin with `inject: ['llm']` can call it), but you must
  pick `provider`/`model` yourself, consume an `AsyncIterable<StreamChunk>` (no
  request/response helper exists), and the call is **not** a session/turn: it will not appear in
  the transcript and is not covered by compaction. There is no `complete`/`generate` convenience
  method anywhere in the package.

---

## 5. THEME TOKENS

`dsh-client-ui-theme/lib/client.js` ships the palettes as literal CSS: static scale on one
selector, alias layer on `body`, dark overrides on `body[data-ds-dark-theme]`:
```css
body{--dsw-alias-bg-base:var(--dsw-static-neutral-bluish-00);--dsw-alias-bg-document-preview:var(--dsw-static-neutral-bluish-100);…}
body[data-ds-dark-theme]{--dsw-alias-bg-base:var(--dsw-static-neutral-bluish-950);…}
```
So tokens are plain custom properties — usable from your panel's CSS with no import.

### Token names actually present (107 `--dsw-alias-*` variables)

```
--dsw-alias-bg-base                        --dsw-alias-label-caption
--dsw-alias-bg-document-preview            --dsw-alias-label-deep-diving
--dsw-alias-bg-document-selection          --dsw-alias-label-deep-diving-shimmer
--dsw-alias-bg-layer-1                     --dsw-alias-label-dimmed
--dsw-alias-bg-layer-2                     --dsw-alias-label-document-preview
--dsw-alias-bg-layer-3                     --dsw-alias-label-primary
--dsw-alias-bg-mask-1                      --dsw-alias-label-primary-bluish
--dsw-alias-bg-mask-2                      --dsw-alias-label-primary-dimmed
--dsw-alias-bg-mask-3                      --dsw-alias-label-primary-foreground
--dsw-alias-bg-mask-drop                   --dsw-alias-label-primary-inverted
--dsw-alias-bg-mask-photo                  --dsw-alias-label-secondary
--dsw-alias-bg-module-platform             --dsw-alias-label-shimmer
--dsw-alias-bg-multi-select                --dsw-alias-label-tertiary
--dsw-alias-bg-overlay                     --dsw-alias-link
--dsw-alias-bg-skeleton                    --dsw-alias-markdown-citation
--dsw-alias-border-inverted                --dsw-alias-markdown-code-block
--dsw-alias-border-inverted2               --dsw-alias-markdown-code-block-banner
--dsw-alias-border-l1                      --dsw-alias-markdown-code-segment-selected
--dsw-alias-border-l2                      --dsw-alias-markdown-code-segment-unselected
--dsw-alias-border-l2-darkmode-thin        --dsw-alias-markdown-inline-code
--dsw-alias-border-l3                      --dsw-alias-markdown-placeholder
--dsw-alias-border-l4                      --dsw-alias-markdown-tag
--dsw-alias-brand-primary                  --dsw-alias-menu-group-header-fill
--dsw-alias-brand-primary-invert           --dsw-alias-menu-icon
--dsw-alias-brand-primary-new-colorprimary-new-color   --dsw-alias-onboarding-accent
--dsw-alias-brand-text                     --dsw-alias-onboarding-card-fill
--dsw-alias-button-contrast-fill           --dsw-alias-onboarding-checkbox-border
--dsw-alias-button-elevated-fill           --dsw-alias-onboarding-secondary-fill
--dsw-alias-button-floating-fill           --dsw-alias-scrollbar-bg-l1
--dsw-alias-button-floating-hover          --dsw-alias-scrollbar-bg-l2
--dsw-alias-button-ghost-active-border     --dsw-alias-scrollbar-hover-l1
--dsw-alias-button-ghost-active-fill       --dsw-alias-scrollbar-hover-l2
--dsw-alias-button-ghost-active-hover      --dsw-alias-settings-card-fill
--dsw-alias-button-info-fill               --dsw-alias-settings-card-stroke
--dsw-alias-button-info-hover              --dsw-alias-state-business-primary
--dsw-alias-button-primary-dimmed          --dsw-alias-state-business-tertiary
--dsw-alias-button-primary-fill            --dsw-alias-state-error-primary
--dsw-alias-button-primary-hover           --dsw-alias-state-error-secondary
--dsw-alias-button-tool-bar-fill           --dsw-alias-state-idle-primary
--dsw-alias-button-tool-bar-fill-invisible --dsw-alias-state-success-primary
--dsw-alias-button-tool-bar-hover          --dsw-alias-state-success-secondary
--dsw-alias-code-diff-added                --dsw-alias-state-success-tertiary
--dsw-alias-code-diff-deleted              --dsw-alias-state-warn-label
--dsw-alias-file-diff-added-bg             --dsw-alias-state-warn-primary
--dsw-alias-file-diff-added-gutter         --dsw-alias-state-warn-secondary
--dsw-alias-file-diff-added-marker         --dsw-alias-state-warn-tertiary
--dsw-alias-file-diff-deleted-bg           --dsw-alias-switch-thumb
--dsw-alias-file-diff-deleted-gutter       --dsw-alias-toast-bg
--dsw-alias-file-diff-deleted-marker       --dsw-alias-toast-label
--dsw-alias-interactive-bg-active          --dsw-alias-tooltip-bg
--dsw-alias-interactive-bg-hover           --dsw-alias-tooltip-key-bg
--dsw-alias-interactive-bg-hover-accent    --dsw-alias-turn-trigger-bg
--dsw-alias-interactive-bg-hover-danger    --dsw-alias-turn-trigger-bg-hover
--dsw-alias-interactive-bg-hover-solid
```
(The `--dsw-alias-brand-primary-new-colorprimary-new-color` name is as it appears in the bundle —
almost certainly a codegen concatenation artifact. **Do not use it.**)

### Radii
```css
--dsw-radius-xs:4px;--dsw-radius-sm:8px;--dsw-radius-md:12px;--dsw-radius-lg:16px;--dsw-radius-xl:20px;--dsw-radius-panel:28px
```

### Fonts
```css
--dsw-font-family:var(--dsh-font-family-text-default);
--dsw-font-family-brand:"Montserrat", var(--dsh-font-family-text-default);
--dsh-font-family-text-default: … "PingFang SC", "Microsoft YaHei", "Helvetica Neue", Helvetica, Arial, sans-serif;
--ds-font-family-code:"SF Mono", "JetBrains Mono", "Fira Code", Consolas, …;
```
Role composites exist per size/weight, e.g.:
```css
--dsw-font-s-14:14px/22px var(--dsw-font-family);   /* + -font-family/-font-size/-font-weight/-font-style/-line-height */
```
Roles available: `xxxs-11`, `xxs-12`, `xs-13`, `s-14`, `base-16`, `m-18`, `l-20`, `xl-24`, plus
`*-strong-*` variants and the `markdown-*` family. Use `--dsw-font-s-14` etc. as a shorthand
`font:` value, or the `-font-size`/`-font-weight`/`-line-height` parts individually.

Brand font files ship in the package (`styles/montserrat-{light,regular,medium}.woff2`,
`styles/brand-font.css`), so `--dsw-font-family-brand` resolves without a network fetch.

Rule of thumb from `references/practices.md`: *"Style with the theme tokens that
`cordis_inspect_query` `Theme` lists (`--dsw-alias-*`); literal colors are for artwork only."*

---

## 6. LOCALE RESOURCES / DISPLAY METADATA

`Plugin Manager and Settings read meta.title and meta.description from exported locale JSON and
an icon from exported resources without activating plugins` (`references/host-plugin.md`).
The reader is `readPluginMeta` in `dsh-app-boot/lib/index.js`:

```js
function readPluginMeta(specifier, parentURL) {
	const packageName = barePackageName(specifier);
	if (packageName === void 0) return void 0;
	try {
		const englishPath = optionalResourcePath(`${specifier}/locale/en.json`, parentURL);
		const dictionaries = englishPath === void 0 ? new Map() : dictionariesOf(englishPath, specifier, parentURL);
		const manifestPath = packageName === specifier ? optionalResourcePath(`${specifier}/package.json`, parentURL) : void 0;
		const manifest = manifestPath === void 0 ? void 0 : readObject(manifestPath);
		const title = localizedText("title", dictionaries, fallbackText(manifest?.name), specifier);
		const description = localizedText("description", dictionaries, fallbackText(manifest?.description), "");
		…
			icon = iconOf(specifier, packageName, parentURL, manifestPath, manifest?.icon);
```
and the parse of each locale file proves the nested `meta` shape:
```js
		const meta = parsed.meta === void 0 ? void 0 : objectOf(parsed.meta, `${file}: meta`);
		dictionaries.set(id, {
			title: textOf(meta?.title, `${file}: meta.title`),
			description: textOf(meta?.description, `${file}: meta.description`)
		});
```
Every `*.json` next to `locale/en.json` becomes a language, filename = language id:
```js
		if (!entry.name.endsWith(".json")) continue;
		const resource = `${specifier}/locale/${entry.name}`;
		const language = entry.name.slice(0, -5);
		if (!LANGUAGE_ID.test(language)) throw new Error(`${resource} must use a language id as its filename`);
```

**Confirmed bundle shape** — exactly what the template ships
(`templates/decoration/locale/en.json`):
```json
{
  "meta": {
    "title": "Composer Decoration",
    "description": "Displays a blue circle below the message composer."
  }
}
```
Required exports/paths: `"./locale/*.json": "./locale/*.json"` and an icon via either
`"./icon": "./icon.svg"` or a top-level `"icon"` in `package.json` (takes priority). Icons:
SVG/PNG/JPEG/WebP, ≤ 256 KiB, inside the package.

**Other namespace keys:** the `meta` namespace is the only one the inventory reads. UI text goes
through your own namespace registered at runtime via `ctx.locale.register(NS, {zh, en})`
(§2c) — that dictionary shape is a flat key→string map per locale, not the `meta` wrapper.
Tool `description` strings are **not** localized (they are literal strings in the definition).

---

## 7. HOST STORAGE

### (a) `ctx.fs` — plain files, the simple option

`dsh-tool-cordis/lib/types/api-catalog.js` (`key: 'fs'`), the members we need:
```js
                signature: 'abstract resolve(path: string, opts?: { cwd?: string; signal?: AbortSignal }): Promise<FsTarget>',
                description: 'Resolve a model/plugin-supplied path into a stable FsTarget. May perform I/O … Relative paths resolve against `opts.cwd`.',
                returns: 'the stable target; the same file yields the same `targetKey`.',
                signature: 'abstract readText(target: FsTarget, signal?: AbortSignal): Promise<string>',
                signature: 'abstract writeText( target: FsTarget, content: string, expected?: FsWriteIntent, signal?: AbortSignal, sandboxPolicy?: SandboxExecutionPolicy, ): Promise<FsWriteOutcome>',
                description: 'Atomically create or replace UTF-8 text. `expected` guards intent and staleness; omission allows unconditional overwrite.',
                signature: 'abstract listDir(target: FsTarget, signal?: AbortSignal): Promise<FsDirEntry[]>',
                description: 'List direct children of a directory in stable name order. Returns resolved child targets plus cheap metadata only; never reads file contents.',
                signature: 'abstract stat(target: FsTarget, signal?: AbortSignal): Promise<FsInfo | undefined>',
                signature: 'abstract lstat(path: string, opts?: { cwd?: string }, signal?: AbortSignal): Promise<FsPathInfo | undefined>',
                signature: 'abstract processPath(target: FsTarget): string',
```
`FsTarget` is deliberately opaque (`targetKey`), so always obtain it from `resolve()`. The
one-flag pattern for "does the file exist" is `stat(...) === undefined` — exactly as
`dsh-tool-present` does:
```js
				const target = await ctx.fs.resolve(file.path, options);
				const info = await ctx.fs.stat(target, exec.signal);
				if (info === void 0) throw new FsError(`Cannot present ${file.path}: file not found. …`, "FS_NOT_FOUND");
```
For a directory anchor without an agent, `workingDirectory` exposes a deployment fallback:
```js
                signature: 'readonly defaultDirectory: string',
                description: 'Absolute deployment fallback for Sessions with no recorded original directory.',
                signature: 'get(session: Session): string',
                signature: 'ensure(agent: Agent, signal?: AbortSignal): Promise<string>',
```

### (b) `ctx.storage` — a domain/KV stack, heavyweight for this job

```js
        key: 'storage',
        summary: 'The storage hub service.',
        description: 'The storage hub service. Backends register under `backend`; data forms mount under their `StorageForms` key and are reached as `ctx.storage.<form>`.',
                signature: 'mount<K extends keyof StorageForms>(form: K, facility: StorageForms[K]): () => void',
```
```js
        key: 'storageDomain',
                signature: 'async open<S extends DomainSpec>(spec: S): Promise<Domain<S>>',
                description: 'Open one declared domain. Steps, each failing the whole call: reject a name that is already open (`already-open`); resolve the backend route (`backend-not-found` …); require its `kv` facet (`facet-unsupported`); open the unit projected from the spec …; load and validate every stored record against the spec\'s zod schemas (`invalid-record` …). … Lifecycle: the CALLER owns the returned handle and closes it via `Domain.close()`.'
```
Using it means authoring a `DomainSpec` (zod schemas + backend route) and owning the
open/close lifecycle — justified for structured multi-record data shared across plugins, but
disproportionate for one JSON document per deck.

### Recommendation

**Use `ctx.fs` with a workspace-relative directory.** Reasons: the deck is a single JSON
document (plus the exported `.pptx`), it should be inspectable/diffable and survivable across
restarts, and a failure is a plain `FS_*` error rather than a schema/version rejection. Key it
by the workspace directory (not the session) so it still works from the root-scoped `main`
panel, which receives no `sessionId` (§2b):

```js
// host half
const DIR = '.branddeck'

async function readJson(ctx, cwd, name) {
  const target = await ctx.fs.resolve(`${DIR}/${name}.json`, { cwd })
  const info = await ctx.fs.stat(target)
  if (info === undefined) return undefined
  return JSON.parse(await ctx.fs.readText(target))
}

async function writeJson(ctx, cwd, name, value) {
  const target = await ctx.fs.resolve(`${DIR}/${name}.json`, { cwd })
  await ctx.fs.writeText(target, JSON.stringify(value, null, 2))  // atomic replace
}
```

---

## 8. BUNDLE / CLIENT BOOT

### `dsh.client` manifest — exact fields

`dsh-client-modules/lib/index.js` validates the declaration; only four fields exist:
```js
function parseDshClient(pkgName, value) {
	if (value === void 0) return void 0;
	if (typeof value !== "object" || value === null) throw new Error(`client-modules: ${pkgName} has a non-object dsh.client declaration`);
	const decl = value;
	if (typeof decl.platform !== "string") throw new Error(`client-modules: ${pkgName} dsh.client.platform must be a string`);
	const inject = optionalStringArray(pkgName, "dsh.client.inject", decl.inject);
	const external = optionalStringArray(pkgName, "dsh.client.external", decl.external);
	if (decl.immediately !== void 0 && typeof decl.immediately !== "boolean") throw new Error(`client-modules: ${pkgName} dsh.client.immediately must be a boolean`);
	return {
		platform: decl.platform,
		...inject !== void 0 ? { inject } : {},
		...external !== void 0 ? { external } : {},
		...decl.immediately !== void 0 ? { immediately: decl.immediately } : {}
	};
}
```
- `platform` — **required**, a string; shipped value `"web"`.
- `inject` — package names whose client rows must be activated first (ordering only).
- `external` — module ids your factory `require`s that another row (or the static table)
  answers; it orders rows and prevents duplicates:
  ```js
   * Order composed rows so every requested dynamic package precedes its
   * consumers. An `external` specifier is either the package row it names
   * (`<pkg>/client` aliases the bare package) or a static-table name that adds no
   * graph edge.
  ```
- `immediately` — boolean; `true` puts the row in the preload batch (`buildBatch`,
  `phase`/`immediately` in `graphRow`).

### Module id **must equal the package name**

> The browser artifact registers a lazy factory whose id equals the package name.
> — `references/ui-plugin.md`

The loader normalizes the `./client` subpath onto the package row, which is why
`id: '@local/x'` and `require('@local/x/client')` agree:
```js
function stripClientSuffix(spec) {
	return spec.endsWith("/client") ? spec.slice(0, -7) : spec;
}
```
The wire artifact is `exports["./client"]`:
```js
function clientExportOf(pkgName, exportsField) {
	…
	const client = exportsField["./client"];
	if (client === void 0) return void 0;
	if (typeof client === "string") return client;
	…
	throw new Error(`client-modules: ${pkgName} exports["./client"] must be a string or an object with a string default`);
```
Loading is lazy and *register-only*; the factory body (including CSS injection) runs at
materialization. A bundle that does not register the expected id fails loudly:
```js
					failures.push(`${url}: loaded without registering "${id}" via __ModuleLoader__.load`);
```
```js
				if (!this.factories.has(id)) throw new Error(`client-modules: bundle ${url} loaded without registering "${id}" via __ModuleLoader__.load`);
```
Baseline `require` targets (safe, no `external` needed): `react`, `react/jsx-runtime`,
`react-dom`, `react-dom/client`, `@deepseek-ai/cordis`, `zod`, `@tanstack/react-virtual`,
`@deepseek-ai/dsh-client-store`, `@deepseek-ai/dsh-client-ui-slots`,
`@deepseek-ai/dsh-client-ui-primitives`, `@deepseek-ai/dsh-client-ui-dockkit`.
(Do **not** `require('@deepseek-ai/dsh-client-ui-primitives')` for widgets — practices.md says
write your own controls; the entry is listed here only to document the table.)

Final template manifest to copy:
```json
{
  "name": "@local/my-decoration",
  "version": "1.0.0",
  "private": true,
  "type": "module",
  "exports": {
    ".": "./index.js",
    "./client": "./client.js",
    "./locale/*.json": "./locale/*.json",
    "./icon": "./icon.svg"
  },
  "files": ["index.js", "client.js", "cordis.patch.yml", "locale/*.json", "icon.svg"],
  "dsh": {
    "bundle": { "patch": "./cordis.patch.yml" },
    "client": { "platform": "web", "immediately": true, "inject": ["@deepseek-ai/dsh-client-ui-conversation"] }
  }
}
```
with `cordis.patch.yml`:
```yaml
- insert:
    - id: my-decoration
      name: '@local/my-decoration'
```

### HMR / rebuild for a linked local directory

The graph revision is derived from **file metadata**, so rewriting `client.js` produces a new rev
without hashing the content:
```js
function artifactRevision(baseline) {
	return framedHash("plugin-artifact", [
		String(baseline.mtimeMs),
		String(baseline.ctimeMs),
		String(baseline.size)
	]);
}
```
```js
	rebuilt(id) {
		const record = this.table.get(id);
		if (record === void 0) return void 0;
		const baseline = this.captureArtifactBaseline(record.meta.clientPath);
		const rev = artifactRevision(baseline);
		if (rev === record.entry.rev) return rev;
		const bundle = readFileSync(record.meta.clientPath);
		record.baseline = baseline;
		record.entry = graphRow(id, rev, record.meta);
		record.bundle = bundle;
		this.composed = this.compose();
		for (const notify of this.rebuildListeners) try {
			notify(id, rev);
```
`rebuilt` is described as *"the HMR watch's registration hook — the only entry point through
which build changes reach the graph"*, and `onRebuilt`/`onGraphChanged` are the subscriptions;
`dsh-client-hmr` is the page-side consumer.

Practically: `install_bundle` with a local directory **links** it (`files` does not filter a
linked directory — `references/host-plugin.md`), so editing `client.js` in place changes the
file the host reads. Whether the running process actually picks it up depends on the HMR watch
being active; the first-party guidance is conservative:

> Installing a new bundle can activate through HMR; replacing an installed package requires
> restart to load a fresh JavaScript module generation. Do not infer updated browser code from
> an unchanged slot id.

**unknown:** the exact watch-installation trigger for a linked workspace directory (which
service calls `rebuilt`) — I did not find the caller in the extracted packages; `dsh-hmr` /
`dsh-client-hmr` are the candidates. Ruled out: it is not in `dsh-host-webserver` (no watcher),
and `rebuilt` has no other caller inside `dsh-client-modules`.

---

## Recommended architecture for the BrandDeck plugin

### Decisions

| Concern | Choice | Why |
|---|---|---|
| Client→Host | `ctx.connection.rpc.handle('/branddeck', handler)` + `ctx.connection.rpc.call('/branddeck', endpoint, payload)` | Supported both halves, automatic authz (owner peer), no codegen, JSON envelope, effect-scoped (§1d) |
| Secondary (agent parity) | one `/deck` slash command via `ctx.commands.register` | Documented in `user-actions.md`; gives the agent the same operation |
| Panel | `main` keyed entry `key:'branddeck'` + `sidebar.panellist` `id:'branddeck'` + `ctx.layout.selectPanel('branddeck')` | Shipped pattern (§2a/2e) |
| Persistence | `ctx.fs` under `<cwd>/.branddeck/` | Simple, durable, inspectable (§7) |
| Deck authoring | agent writes JSON through tools; panel edits the same JSON | Avoids the streaming-only LLM seam (§4b) |
| Styling | `--dsw-alias-*` + `--dsw-radius-*` + `--dsw-font-*` only | Theme-consistent, upgrade-safe (§5) |

### Package layout

```
branddeck/
  package.json
  cordis.patch.yml
  index.js          # host half
  client.js         # browser half (window.__ModuleLoader__)
  locale/en.json    # {"meta":{"title","description"}}
  locale/zh.json
  icon.svg
```

### `package.json`

```json
{
  "name": "@local/branddeck",
  "version": "1.0.0",
  "private": true,
  "type": "module",
  "exports": {
    ".": "./index.js",
    "./client": "./client.js",
    "./locale/*.json": "./locale/*.json",
    "./icon": "./icon.svg"
  },
  "files": ["index.js", "client.js", "cordis.patch.yml", "locale/*.json", "icon.svg"],
  "dsh": {
    "bundle": { "patch": "./cordis.patch.yml" },
    "client": {
      "platform": "web",
      "immediately": true,
      "inject": ["@deepseek-ai/dsh-client-ui-layout", "@deepseek-ai/dsh-client-ui-sidebar"]
    }
  }
}
```

### `cordis.patch.yml`

```yaml
- insert:
    - id: branddeck
      name: '@local/branddeck'
      config: {}
```

### Host half `index.js` (skeleton)

```js
/** Host half of the BrandDeck bundle: deck store, pptx export, agent tools, RPC channel. */
import { defineTool } from '@deepseek-ai/dsh-tools'   // resolves from the dsh installation

const name = 'branddeck'
const inject = ['connection', 'tools', 'fs', 'workingDirectory', 'commands', 'systemPrompt']

const DIR = '.branddeck'

// Placeholders owned by the deck engine (not part of the DSH API surface):
//   readThemeTokens(templatePath)  -> { colors, fonts, logoDataUrl }
//   exportPptx(deck, { outFile, cwd })
//   emptyDeck()
// Keep them plain functions so the tools, the RPC channel, and unit tests all share one path.

export function apply(ctx, config) {
  const cwdOf = () => ctx.workingDirectory.defaultDirectory

  const deckPath = async (cwd, id) => ctx.fs.resolve(`${DIR}/${id}.json`, { cwd })
  const readDeck = async (id) => {
    const target = await deckPath(cwdOf(), id)
    if (await ctx.fs.stat(target) === undefined) return undefined
    return JSON.parse(await ctx.fs.readText(target))
  }
  const writeDeck = async (id, value) => {
    const target = await deckPath(cwdOf(), id)
    await ctx.fs.writeText(target, JSON.stringify(value, null, 2)) // atomic, unconditional
    return value
  }

  // ---- one operation, three callers: RPC (panel), tool (agent), command (user) ----
  const operations = {
    'theme.read': async (_payload) => {
      // parse .pptx template -> { colors, fonts, logoDataUrl }  (pptxgenjs/zip on host)
      return { ok: true, value: await readThemeTokens(config.templatePath) }
    },
    'deck.get': async ({ id }) => ({ ok: true, value: (await readDeck(id)) ?? emptyDeck() }),
    'deck.put': async ({ id, deck }) => ({ ok: true, value: await writeDeck(id, deck) }),
    'deck.export': async ({ id, outFile }) => {
      const deck = await readDeck(id)
      const path = await exportPptx(deck, { outFile, cwd: cwdOf() })
      return { ok: true, value: { path } }
    }
  }

  // RPC channel: admission (Host fence + cookie) is applied by connection.register(),
  // which already wraps the route in owner.effect(...) — no extra ctx.effect needed.
  // Result MUST be { ok: true, value } or { ok: false, error: { code, message, details } }.
  const disposeRpc = ctx.connection.rpc.handle('/branddeck', async (endpoint, payload, signal, peer) => {
    if (peer.principal.kind !== 'owner') {
      return { ok: false, error: { code: 'branddeck/forbidden', message: 'owner only', details: {} } }
    }
    const operation = operations[endpoint]
    if (operation === undefined) {
      return { ok: false, error: { code: 'branddeck/unknown-endpoint', message: endpoint, details: {} } }
    }
    try {
      return await operation(payload ?? {}, signal)
    } catch (error) {
      return { ok: false, error: { code: 'branddeck/failed', message: String(error?.message ?? error), details: {} } }
    }
  })

  // Agent tool over the same operations. defineTool compiles the compact spec to
  // JSON Schema AND installs argument validation — pass its result to register().
  ctx.tools.register(defineTool({
    name: 'branddeck_deck',
    description: 'Read or write the BrandDeck slide deck for the current workspace, or export it to .pptx.',
    parameters: {
      action: { type: 'string', required: true, enum: ['get', 'put', 'export'], description: 'Operation to perform.' },
      id: { type: 'string', required: true, description: 'Deck id, e.g. "main".' },
      deck: { type: 'object', description: 'Full deck JSON, required for action=put.' },
      outFile: { type: 'string', description: 'Output .pptx path, for action=export.' }
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_args, value) => [{ type: 'text', text: JSON.stringify(value) }]
    },
    async execute(args, exec) {
      const endpoint = args.action === 'get' ? 'deck.get' : args.action === 'put' ? 'deck.put' : 'deck.export'
      const result = await operations[endpoint](args, exec.signal)
      if (result.ok !== true) throw new Error(result.error.message)
      return result.value
    }
  }))

  // Slash command so the agent and the user reach the same operation
  ctx.effect(() => ctx.commands.register({
    name: 'deck',
    description: 'Export the BrandDeck deck to .pptx',
    handler: async (invocation) => {
      const result = await operations['deck.export']({ id: 'main' }, invocation.signal)
      return result.ok
        ? { kind: 'success', text: `Exported ${result.value.path}` }
        : { kind: 'error', text: result.error.message }
    }
  }), 'branddeck: /deck command')

  // Prompt section: tell the agent the deck JSON contract
  ctx.effect(() => ctx.systemPrompt.section({
    name: 'branddeck:contract',
    order: 2650,
    text: 'Decks live in .branddeck/<id>.json with shape {slides:[{title,body,notes}]}. Use branddeck_deck to read/write/export.'
  }), 'branddeck: prompt section')
}
```
Notes: `CommandDefinition` accepts `{definitionId?, name, description, handler}` — `definitionId`
is optional, so no `@deepseek-ai/dsh-commands` import is needed
(`dsh-commands/lib/index.js`: `...definition.definitionId === void 0 ? {} : { definitionId: definition.definitionId }, name: …, description: …, handler: definition.handler`).

### Client half `client.js` (skeleton, including the transport call)

```js
window.__ModuleLoader__.load({
  id: '@local/branddeck',                       // MUST equal the package name
  factory(require) {
    const React = require('react')
    const h = React.createElement
    const NS = 'branddeck'
    const PANEL_ID = 'branddeck'

    // ---- transport: generic Connection RPC (auth handled by the host registry) ----
    async function call(ctx, endpoint, payload) {
      const result = await ctx.connection.rpc.call('/branddeck', endpoint, payload)
      if (result.ok !== true) throw new Error(result.error?.message ?? 'branddeck: call failed')
      return result.value
    }

    // ---- 16:9 slide card, styled with real theme tokens ----
    function SlideCard({ slide, index, onEdit }) {
      return h('div', {
        style: {
          aspectRatio: '16 / 9',
          background: 'var(--dsw-alias-bg-layer-1)',
          border: '1px solid var(--dsw-alias-border-l2)',
          borderRadius: 'var(--dsw-radius-md)',
          padding: '16px',
          font: 'var(--dsw-font-s-14)',
          color: 'var(--dsw-alias-label-primary)'
        }
      },
        h('div', { style: { font: 'var(--dsw-font-l-20)', color: 'var(--dsw-alias-label-primary)', marginBottom: 8 } },
          slide.title || `Slide ${index + 1}`),
        h('div', { style: { color: 'var(--dsw-alias-label-secondary)', whiteSpace: 'pre-wrap' } }, slide.body || ''),
        h('button', {
          onClick: () => onEdit(index),
          style: {
            marginTop: 12,
            background: 'var(--dsw-alias-button-primary-fill)',
            color: 'var(--dsw-alias-label-primary-inverted)',
            borderRadius: 'var(--dsw-radius-sm)',
            border: 'none',
            padding: '6px 12px',
            cursor: 'pointer'
          }
        }, 'Edit')
      )
    }

    function BrandDeckPanel(props) {
      const [deck, setDeck] = React.useState(null)
      const [error, setError] = React.useState(null)
      const [busy, setBusy] = React.useState(false)

      React.useEffect(() => {
        let cancelled = false
        props.api.getDeck().then((value) => { if (!cancelled) setDeck(value) }, (reason) => { if (!cancelled) setError(String(reason)) })
        return () => { cancelled = true }
      }, [])

      const update = (index, field, text) => {
        setDeck((current) => {
          const slides = current.slides.slice()
          slides[index] = { ...slides[index], [field]: text }
          return { ...current, slides }
        })
      }

      const save = async () => {
        setBusy(true)
        try { await props.api.putDeck(deck) } catch (reason) { setError(String(reason)) } finally { setBusy(false) }
      }

      return h('div', { style: { padding: 16, background: 'var(--dsw-alias-bg-base)', height: '100%', overflow: 'auto' } },
        h('div', { style: { display: 'flex', gap: 8, marginBottom: 12 } },
          h('button', { onClick: save, disabled: busy, style: buttonStyle }, busy ? 'Saving…' : 'Save'),
          h('button', { onClick: () => props.api.exportDeck(), style: buttonStyle }, 'Export .pptx')
        ),
        error === null ? null : h('div', { style: { color: 'var(--dsw-alias-state-error-primary)' } }, error),
        h('div', { style: { display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 12 } },
          (deck?.slides ?? []).map((slide, index) => h(SlideCard, {
            key: index,
            slide,
            index,
            onEdit: () => undefined // drag/edit affordances go here
          }))
        )
      )
    }

    const buttonStyle = {
      background: 'var(--dsw-alias-button-elevated-fill)',
      color: 'var(--dsw-alias-label-primary)',
      border: '1px solid var(--dsw-alias-border-l1)',
      borderRadius: 'var(--dsw-radius-sm)',
      padding: '6px 12px',
      cursor: 'pointer',
      font: 'var(--dsw-font-s-14)'
    }

    // ---- sidebar icon: switches the main column to this panel ----
    function BrandDeckIcon(props) {
      return h('button', {
        onClick: () => props.ctx.layout.selectPanel(PANEL_ID),
        title: props.t('panel'),
        style: {
          width: props.size, height: props.size, padding: 0,
          border: 'none', borderRadius: 'var(--dsw-radius-sm)',
          background: props.active ? 'var(--dsw-alias-interactive-bg-active)' : 'transparent',
          color: props.active ? 'var(--dsw-alias-label-primary)' : 'var(--dsw-alias-menu-icon)',
          cursor: 'pointer', display: 'grid', placeItems: 'center'
        }
      }, h('svg', { viewBox: '0 0 24 24', width: props.size * 0.6, height: props.size * 0.6, 'aria-hidden': true },
        h('rect', { x: 3, y: 5, width: 18, height: 12, rx: 2, fill: 'none', stroke: 'currentColor', strokeWidth: 1.6 }),
        h('path', { d: 'M12 5v12M3 9h18', stroke: 'currentColor', strokeWidth: 1.2 })))
    }

    return {
      inject: ['slots', 'locale', 'connection', 'layout'],
      apply(ctx) {
        ctx.effect(() => ctx.locale.register(NS, {
          en: { panel: 'BrandDeck', save: 'Save', export: 'Export .pptx' },
          zh: { panel: '幻灯片', save: '保存', export: '导出 .pptx' }
        }), 'branddeck: dictionaries')
        const t = ctx.locale.bind(NS)

        const api = {
          getDeck: () => call(ctx, 'deck.get', { id: 'main' }),
          putDeck: (deck) => call(ctx, 'deck.put', { id: 'main', deck }),
          exportDeck: () => call(ctx, 'deck.export', { id: 'main' })
        }

        // keyed main panel — `key`, not `id`; no sessionId prop exists for this slot
        ctx.slots.inject('main', () => ctx.slots.register({
          name: 'main',
          key: PANEL_ID,
          locale: NS,
          inject: () => ({ api, t })
        }, BrandDeckPanel))

        // list slot icon — `id`, not `key`; label is a thunk so it follows the locale
        ctx.slots.inject('sidebar.panellist', () => ctx.slots.register({
          name: 'sidebar.panellist',
          id: PANEL_ID,
          order: 20,
          locale: NS,
          label: () => t('panel'),
          inject: () => ({ ctx, t })
        }, BrandDeckIcon))
      }
    }
  }
})
```

---

## Blockers / open items

1. **Slot `BaseOptions` is not fully documented in the extracted artifacts.** `ui-slots` is not
   installed, so beyond `name`/`key`/`id`/`order`/`label`/`locale`/`inject` (all confirmed by
   shipped registrations) the remaining fields are **unknown**. Ruled out: it is not in
   `dsh-cordis-client-runner`'s catalog (which only enumerates per-slot options), and no `.d.ts`
   for the package exists on disk.
2. **Whether the `locale` slot option injects a `t` prop** into the component is **unknown**;
   the sketch passes `t` explicitly through `inject` instead of relying on it.
3. **`GenerateOptions` / `LlmCallConfig` full shape is unknown** (types erased at runtime). Only
   `provider`, `model`, `messages`, `signal` are observed in the implementation, and there is no
   `complete()` helper — a direct LLM pipeline must consume `ctx.llm.stream()`.
4. **HMR trigger for a linked workspace directory is unknown.** `rebuilt(id)` is the documented
   entry point and is mtime/size-driven, but I could not find its caller in the extracted tree
   (`dsh-hmr` / `dsh-client-hmr` are the candidates). The safe assumption from the first-party
   docs is: a *new* bundle can activate via HMR, a *replaced* package needs a restart.
5. **The active session id for a root-scoped `main` panel is not directly available** — the slot
   gives `useSessions`/`usePanelInfo` hooks but no `sessionId` prop. The design above sidesteps
   this by scoping decks to the workspace directory. If per-session decks are truly required, the
   session key must come from `ctx.sessions`/`ctx.uiWorkspace`, whose exact "active session"
   selector I did **not** verify.
6. `cordis_inspect_query` was not usable in this session (`input must be an object` rejected),
   so every conclusion above comes from source reading only. The live `Slots.listSubTree` /
   `Theme` responses would be the cheapest way to double-check items 1 and 2 once available.

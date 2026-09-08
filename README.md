# Teams desktop userscript loader proof of concept

This is a sidecar injector. It does not modify the Teams MSIX/app bundle, its
signature, or the downloaded Teams JavaScript. It attaches to the localhost
Chrome DevTools Protocol (CDP) endpoint exposed by the embedded web runtime.

## 1. Expose Teams on a local CDP port

Use a non-default port to reduce collisions with browsers and other WebView2
applications.

### Windows

Preferred, scoped to one Teams launch:

```powershell
$package = Get-ChildItem "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\Repository\Packages" |
  Where-Object { $_.PSChildName -like "MSTeams_*8wekyb3d8bbwe" } |
  Sort-Object PSChildName -Descending |
  Select-Object -First 1

Stop-Process -Name ms-teams -Force -ErrorAction SilentlyContinue
Start-Process (Join-Path $package.GetValue("PackageRootFolder") "ms-teams.exe") `
  -ArgumentList "--remote-debugging-port=9223"
```

If that Teams build ignores the executable argument, use the broader fallback:

```powershell
[Environment]::SetEnvironmentVariable(
  "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS",
  "--remote-debugging-port=9223",
  "User"
)
```

Fully quit and restart Teams after setting it. This fallback affects every
WebView2 app launched by the user until the variable is removed.

### macOS

GUI applications do not inherit values exported in a terminal:

```bash
launchctl setenv WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS \
  "--remote-debugging-port=9223"
```

Fully quit and restart Teams.

### Verify

One of these should return a JSON target list:

```bash
curl http://127.0.0.1:9223/json/list
curl "http://[::1]:9223/json/list"
```

## 2. Run the loader

Node 22 or newer is recommended:

```bash
node teams-userscript-loader.mjs \
  --port 9223 \
  --scripts ./userscripts
```

### Run automatically on macOS

Install and start the two per-user launch agents:

```bash
npm run service:install
```

The installer derives the repository path, current Node executable, home
directory, and GUI user ID. It configures the Teams CDP environment at login and
keeps the userscript loader running. Fully quit and reopen Teams after
installation.

Logs are written to:

```text
~/Library/Logs/teams-user-scripts/loader.log
~/Library/Logs/teams-user-scripts/loader-error.log
```

To stop and remove both launch agents:

```bash
npm run service:uninstall
```

The loader:

- discovers all current Teams page targets;
- runs matching scripts in already-loaded documents;
- uses `Page.addScriptToEvaluateOnNewDocument` for reloads and new frames;
- watches the userscript directory and reinjects edited scripts;
- discovers new Teams windows, such as meeting and calendar windows.

The included `userscripts/teams-vimium.user.js` adds keyboard-driven Teams
navigation. Press `?` in Teams to see its key map. It includes:

- Vim-style scrolling and pane selection;
- spatial `h/j/k/l` navigation between Calendar events, with `Enter` to open
  the focused event and `[[`/`]]` to change calendar periods;
- previous/next screen navigation;
- modal dismissal and compose-box insert mode;
- non-destructive full-page find highlighting;
- short generated hints for visible interactive elements, using home-row keys
  first and the rest of the alphabet when needed to keep busy screens at two
  keystrokes;
- a searchable vomnibar that includes the full virtualized chat rail, visible
  Teams controls, or every event in the current Calendar period.

The script deliberately does not intercept regular typing in inputs, textareas,
comboboxes, or content-editable message composers. `Escape` is the sole
exception: it returns to normal mode and removes focus.

`userscripts/teams-cs-dev-toggle.user.js` adds a `cs-dev` toggle beside the
title-bar ring badge. Enabling it sets
`calling.conversationServiceUrlOverride.url` in `tmp.settings` to
`https://api.conv-dev.skype.net/conv/`; disabling it removes the override.
Either action uses the ring popover's **Restart (also exits container)** control
so the setting takes effect.

`userscripts/teams-user-extensions.user.js` adds **User extensions** to the
Settings and more menu. Its modal lists matching userscripts and lets each one
be enabled or disabled. The choice is stored in `teams.userscripts.disabled`
and applied by restarting Teams. Scripts can expose a settings modal by
registering a function in `globalThis.__teamsUserscriptSettings` under their
metadata name. The extension manager shows a **Settings** link for registered
scripts. The extension manager itself cannot be disabled.

`userscripts/meeting-meter.user.js` displays an estimated running meeting cost
next to the call timer. It counts participants from the meeting roster control
and uses a `$200,000 USD` fully loaded annual cost per participant, spread over
2,080 working hours per year.

Supported metadata is intentionally small:

```javascript
// ==UserScript==
// @name          My script
// @match         https://teams.cloud.microsoft/*
// @exclude-match https://teams.cloud.microsoft/v2/blocked/*
// @run-at        document-start
// ==/UserScript==
```

`@match`, `@include`, `@exclude-match`, `@exclude`, `@run-at`, and
`@toggleable` are supported. Set `@toggleable false` for infrastructure scripts
that must remain enabled. The matcher is wildcard-based, not a complete
implementation of the browser match-pattern grammar. There are no `GM_*` APIs,
extension background workers, extension storage, or privileged network
requests.

## 3. Cleanup

Stop the loader, restart Teams without the debug argument, and remove any
persistent environment variable:

```powershell
[Environment]::SetEnvironmentVariable(
  "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS",
  $null,
  "User"
)
```

```bash
launchctl unsetenv WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS
```

## Security boundary

CDP has no authentication. Any process able to reach the endpoint can inspect
and control the signed-in Teams web content. Keep it bound to loopback, never
forward or proxy the port, and do not leave debugging enabled when it is not in
use. Only load scripts you trust.

## Why not patch the package

Editing the Windows MSIX or macOS app bundle invalidates code-signing and package
integrity, needs elevated/developer installation paths, and is overwritten by
Teams updates. Rewriting downloaded web assets through a proxy adds TLS,
service-worker, CSP, cache, and source-map problems. The sidecar preserves the
package and moves the unstable part to a small local tool.

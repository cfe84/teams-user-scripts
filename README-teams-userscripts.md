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

The loader:

- discovers all current Teams page targets;
- runs matching scripts in already-loaded documents;
- uses `Page.addScriptToEvaluateOnNewDocument` for reloads and new frames;
- watches the userscript directory and reinjects edited scripts;
- discovers new Teams windows, such as meeting and calendar windows.

Supported metadata is intentionally small:

```javascript
// ==UserScript==
// @name          My script
// @match         https://teams.cloud.microsoft/*
// @exclude-match https://teams.cloud.microsoft/v2/blocked/*
// @run-at        document-start
// ==/UserScript==
```

`@match`, `@include`, `@exclude-match`, `@exclude`, and `@run-at` are supported.
The matcher is wildcard-based, not a complete implementation of the browser
match-pattern grammar. There are no `GM_*` APIs, extension background workers,
extension storage, or privileged network requests.

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

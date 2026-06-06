---
name: coffee-auto-button
description: Coffee AutoButton WPF app maintenance guide. Use this skill when working in C:\work\CoffeeWinApp on non-intrusive click behavior, dedicated browser CDP recognition, WPF UI/settings, .NET 10 builds, publish scripts, MSI installer generation, or version updates.
---

# Coffee AutoButton

## Scope

Use this skill for Coffee AutoButton work in `C:\work\CoffeeWinApp`. The project is a Windows WPF app that automates key input and fixed-position clicks while prioritizing non-interference with the user's cursor and active window.

## Current Core Rules

- Treat Chrome, Edge, Chromium, Brave, Vivaldi, Opera, and Electron-like processes as Chromium targets.
- Chromium targets must use the dedicated browser CDP path on port `9223`.
- Do not add `PostMessage` or physical-click fallback for Chromium targets.
- Do not call `window.focus`, simulate focus, set Topmost, or otherwise steal activation for browser clicking.
- Require successful `BrowserClickTarget` recognition before sending a Chromium click.
- Non-Chromium targets may use Win32 `PostMessage` clicking.
- Physical click is an explicit compatibility path only because it moves the cursor.

## Dedicated Browser Behavior

- Launch Chrome with an isolated profile under `%APPDATA%\CoffeeAutoButton\dedicated-browser-profile`.
- Use the URL from the dedicated browser URL setting when opening the browser.
- Position capture should refresh browser recognition and show `専用ブラウザ認識: CDP:9223 / ...` when successful.
- If Chrome is not recognized as the dedicated browser, stop and explain the dedicated-browser requirement.

## Versioning

When changing release behavior or producing a new installer, keep these aligned:

- `CoffeeAutoButton/CoffeeAutoButton.csproj` `Version`
- `AssemblyVersion`
- `FileVersion`
- `tools/build-msi.ps1` default version
- UI version display
- README/DESIGN version notes when relevant

## Build And Release Checks

Common build command:

```powershell
dotnet build .\CoffeeAutoButton.sln
```

Common MSI command:

```powershell
.\tools\build-msi.ps1 -PublishPath "publish\CoffeeAutoButton-<version>"
```

If the default publish folder is locked by a running app, publish to a versioned output folder instead of killing the app unless the user approves.

After MSI work, verify:

- `installer\out\CoffeeAutoButtonSetup.msi` exists.
- Installed exe is under `%LOCALAPPDATA%\CoffeeAutoButton`.
- Windows installed-app DisplayVersion matches the exe version.
- Start menu shortcut is created as `Coffee AutoButton`.

## Documentation

After behavior changes, update:

- `README.md` for user-facing concept, usage, build, and installer notes.
- `DESIGN.md` for implementation rules and current design constraints.
- `CHECKLIST.md` when manual verification steps change.
- This skill when the project rules or release workflow change.

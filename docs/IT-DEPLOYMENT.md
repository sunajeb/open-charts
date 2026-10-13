# Deploying open-charts: a guide for IT

This page is for IT admins who deploy open-charts to Windows PCs and Macs. It says what the
installers put where, what talks to the network, how to deploy and remove it silently, and which
Office and macOS settings it needs. How the installers are built and signed is in
[release.md](release.md). Windows internals are in [windows.md](windows.md).

There are two add-ins. You can deploy either or both; charts made by one can be edited by the other.

| | Macro add-in (ribbon tab) | Office.js add-in (task pane) |
|---|---|---|
| Runs on | PowerPoint for Windows (Office 2016 or newer, Microsoft 365), PowerPoint for Mac (version 16) | PowerPoint for Windows, Mac and the web, wherever Office add-ins run |
| Ships as | `open-charts-<version>.msi`, `open-charts-<version>-machine.msi`, `open-charts-<version>.pkg` | static web files plus `manifest.xml` |
| Needs | macros allowed for one trusted file; on a Mac, one Automation permission | nothing on the computer; a https host for the files |
| Deploy with | Intune, SCCM, Jamf | Microsoft 365 admin center, [OFFICEJS-HOSTING.md](OFFICEJS-HOSTING.md) |

**Not done yet, said plainly:**

- **Code signing is pending.** Today's release files are unsigned: `open-charts.exe`, the MSIs, the
  `.pkg` and the VBA project inside `open-charts.ppam`. The signing steps are ready
  ([release.md](release.md#signing-what-the-owner-needs)) and wait for the certificates. Until then
  you cannot allow open-charts by publisher, only by path or hash.
- **There is no machine-wide Mac package.** The `.pkg` installs for one user at a time.

## What gets installed where

### Windows, per user: `open-charts-<version>.msi`

No admin rights; installs for the user who runs it.

| What | Where |
|---|---|
| The engine (Node and the chart engine in one exe, about 80 MB, x64) | `%LOCALAPPDATA%\open-charts\open-charts.exe` |
| Its callback prefix, uninstall script | `%LOCALAPPDATA%\open-charts\addin.json`, `uninstall.ps1` |
| The add-in | `%APPDATA%\Microsoft\AddIns\open-charts.ppam` |
| Load with PowerPoint | `HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\open-charts` (`Path`, `AutoLoad`=1) |
| Trusted location for the add-ins folder | `HKCU\Software\Microsoft\Office\16.0\PowerPoint\Security\Trusted Locations\open-charts` (`Path`, `Description`) |
| Install folder, for the uninstall | `HKCU\Software\open-charts` (`engineDir`) |

Written later, while it runs, all under `%LOCALAPPDATA%\open-charts`:

- `token`: the local server's secret (below);
- `addin.log`: results and errors;
- `settings.json`, and `update.json` if the update check is on;
- `res\<hash>\`: resources the exe unpacks on first start;
- `edge-popup\`: the profile of the Edit data window, a separate Microsoft Edge profile.

Uninstalling the MSI removes all of it, including the files the engine wrote.

### Windows, per machine: `open-charts-<version>-machine.msi`

For Intune and SCCM, run as SYSTEM.

- Installs `open-charts.exe`, `open-charts.ppam`, `addin.json`, `install.ps1` and `uninstall.ps1` to
  `%ProgramFiles%\open-charts`. It is a 64-bit package.
- Adds an Active Setup entry:
  `HKLM\Software\Microsoft\Active Setup\Installed Components\open-charts`. Its `StubPath` runs
  `powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File "<install dir>\install.ps1" -SkipEngineCheck`.
- At each user's next sign-in, Windows runs that step once, and again when the version goes up.
  It puts the per-user files and keys from the table above in place. It also creates `token`.
  It adds the trusted location only if the add-ins folder is not trusted already. It does not
  write `HKCU\Software\open-charts`.
- A user who is signed in during the install gets the add-in after signing out and in again.

### Mac: `open-charts-<version>.pkg`

The package installs into the user's home folder only. It asks for no password. Installing it
as root to the system volume is refused.

| What | Where (`~` is the user's home folder) |
|---|---|
| The engine (one universal executable, no Node needed) | `~/Library/Application Support/open-charts/engine/open-charts` |
| The Edit data window (`org.open-charts.popup`) | `~/Library/Application Support/open-charts/OpenChartsPopup.app` |
| Package copy of the add-in and the uninstaller | `~/Library/Application Support/open-charts/package/` |
| The add-in | `~/Library/Application Support/Microsoft/open-charts/open-charts.ppam` |
| The bridge script PowerPoint runs | `~/Library/Application Scripts/com.microsoft.Powerpoint/open-charts.scpt` |
| The `open-charts` command | `~/Library/Application Support/open-charts/bin/open-charts`, linked into `~/.local/bin`, `~/bin`, `/opt/homebrew/bin` or `/usr/local/bin` only if one is on the user's `PATH` and writable |
| The uninstaller | `~/Applications/Uninstall open-charts.app` |
| A short-lived LaunchAgent (below) | `~/Library/LaunchAgents/io.github.open-charts.register.plist` |
| The package receipt | package id `io.github.open-charts.addin`, in the user's home volume |

Also written under `~/Library/Application Support/open-charts`: `token` (mode 600), `addin.json`,
`addin.log`, `install.log`, `settings.json`, and `update.json` if the update check is on.

The installer registers the add-in in PowerPoint without starting it. If PowerPoint is running,
it registers it right away. If not, the LaunchAgent runs `pgrep` every 30 seconds. Once
PowerPoint runs, the agent registers the add-in and removes itself. After 10 failed attempts it
gives up and writes that to `addin.log`.

## Network use and data

- **The local server.** The engine runs a small web server on `127.0.0.1:4174`. The add-in starts
  it when needed, and it stops after an hour without requests. It listens on the loopback
  address only, so other computers cannot reach it.
- **The token.** The server answers only callers that send the install's token. The token is 32
  random bytes in the `token` file, which only that user can read. It also refuses requests that
  are not addressed to `127.0.0.1` or `localhost`, and requests sent from other web sites. So a web
  page in the user's browser cannot drive PowerPoint or read the charts being edited.
- **Proxies.** On Windows, the add-in's HTTP calls to the server bypass any WinHTTP proxy. On a Mac,
  the bridge calls it with `curl`.
- **The update check is off by default.** When a user turns it on (`open-charts addin update on`),
  the engine asks `https://api.github.com/repos/sunajeb/open-charts/releases/latest` at most once a
  day. The request carries only `User-Agent: open-charts/<version>`. It only reports a new version
  in the log. See [Updates](#updates) for how to force it off.
- **Nothing else leaves the computer.** There is no telemetry and no account. Charts, data and
  fonts are read and laid out on the computer. The Edit data window is a local page served by the
  local server.

## Windows

### Intune

**Per machine (recommended):**

- Add `open-charts-<version>-machine.msi` as a *line-of-business app*, or wrap it as a Win32 app
  (`.intunewin`).
- Install command: `msiexec /i open-charts-<version>-machine.msi /quiet`
- Uninstall command: `msiexec /x {product code} /quiet`
- Install behavior: *System*. Assign to device groups.
- Detection rule: *MSI*, with the package's product code (Intune reads it from the MSI). Each
  version has its own product code; the upgrade code stays the same, so a newer MSI replaces the
  older one. Or detect the file `%ProgramFiles%\open-charts\open-charts.exe`.

**Per user:**

- Assign `open-charts-<version>.msi` to user groups, install behavior *User*.
- Install command: `msiexec /i open-charts-<version>.msi /quiet`. It needs no admin rights.
- Detection rule: *MSI* product code.

### SCCM (Configuration Manager)

Create an application with a *Windows Installer (\*.msi file)* deployment type:

- **Machine MSI:** installation behavior *Install for system*. Install with
  `msiexec /i "open-charts-<version>-machine.msi" /qn` and uninstall with
  `msiexec /x {product code} /qn`. Detection: the MSI product code.
- **Per-user MSI:** installation behavior *Install for user*, the same commands with
  `open-charts-<version>.msi`.

### Uninstall

- **Per-user MSI:** removes everything it installed and everything the engine wrote.
- **Machine MSI:** removes `%ProgramFiles%\open-charts` and the Active Setup entry. Each user's
  own copy stays until `uninstall.ps1` runs as that user, for example from an Intune script set to
  run in the user's context:

  ```powershell
  powershell -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\open-charts\uninstall.ps1"
  ```

  It stops the engine, removes the two `HKCU` keys, `open-charts.ppam` and
  `%LOCALAPPDATA%\open-charts`.
- Charts already in presentations stay as they are. They are plain PowerPoint shapes.

### Macros and Trust Center policy

open-charts is a macro add-in (`.ppam`). PowerPoint runs its VBA only if policy allows macros in
that file.

**What the installers do:**

- They load the add-in from `%APPDATA%\Microsoft\AddIns\`, PowerPoint's own add-ins folder.
- They add that folder as a trusted location under `HKCU`, unless it is trusted already.
- `install.ps1` also removes the "downloaded from the internet" mark from the files.

**What works today (unsigned VBA): a trusted location.** If your policy manages trusted
locations, add the user's `%APPDATA%\Microsoft\AddIns\` folder there. In Group Policy (Office
ADMX templates), or in the Intune settings catalog under the same names:

- *Microsoft PowerPoint 2016 › PowerPoint Options › Security › Trust Center › Trusted Locations*:
  set **Trusted Location #1** (or the next free number) to that folder. Keep **Disable all trusted
  locations** off.
- *Microsoft Office 2016 › Security Settings › Trust Center*: **Allow mix of policy and user
  locations** lets the `HKCU` location the installer writes count beside your policy ones.

**VBA Macro Notification Settings** (*Microsoft PowerPoint 2016 › PowerPoint Options › Security ›
Trust Center*; the registry value `VBAWarnings` under
`HKCU\Software\Policies\Microsoft\Office\16.0\PowerPoint\Security`):

| Setting | `VBAWarnings` | open-charts |
|---|---|---|
| Enable all macros | 1 | runs |
| Disable all with notification | 2 | runs from the trusted location |
| Disable all except digitally signed macros | 3 | runs from the trusted location; once the VBA project is signed, also by publisher |
| Disable all without notification | 4 | `install.ps1` warns that this blocks the add-in. Microsoft documents that files in trusted locations still run under it; this is not tested with open-charts yet, so try it in a pilot. |

**Once the VBA project is signed** (task C3, not done yet), you can instead deploy the signing
certificate to the users' *Trusted Publishers* store and use *Disable all except digitally signed
macros* without a trusted location. Until then, policies that accept only signed add-ins block it.
That includes **Require that application add-ins are signed by Trusted Publisher**, and
**Disable all application add-ins** blocks it too.

**Application control (AppLocker, WDAC):** allow `%LOCALAPPDATA%\open-charts\open-charts.exe`. The
VBA starts it, and it starts `powershell.exe` (Windows PowerShell 5.1) for Apply and Microsoft Edge
for the Edit data window. Until the exe is signed, allow it by path or file hash. The hash changes
with each version. `install.ps1` stops with a message naming the file if it cannot start it.

### Living with other PowerPoint add-ins

- **Ribbon:** its own tab, **open-charts**, before Home (`tabOpenCharts`). It changes no other tab.
  On Windows it adds its commands to the right-click menu of its own charts only.
- **Macros:** every callback has a per-build prefix (`OC` and 24 hex digits), so no other add-in
  has the same macro names.
- **Charts:** open-charts charts are grouped shapes tagged `OC_*`. It acts only on shapes tagged
  `OC_V`, so it never changes another add-in's charts or elements.
- **Events:** it listens to PowerPoint's selection change and does nothing for shapes that are not
  its own charts.
- **Files:** no shared files, registry keys or COM registration. Its files and keys are only the
  ones listed above.
- **Keyboard shortcuts:** it adds none to PowerPoint. The shortcuts in its Edit data window work
  only in that window.

## Mac

### Deploying the `.pkg` (Jamf, Intune for Mac, other MDM)

The package installs per user, so run it **as the signed-in user**, not as root. MDM "install
package" actions that install to `/` as root are refused, because the package does not offer the
system volume.

Jamf policy script (runs as root, installs as the console user):

```sh
user=$(stat -f %Su /dev/console); uid=$(id -u "$user")
launchctl asuser "$uid" sudo -u "$user" installer -pkg /path/open-charts-<version>.pkg -target CurrentUserHomeDirectory
```

Or put the package in **Self Service**, where the user clicks Install. In Intune for Mac, use a
shell script set to *run as signed-in user* that runs `installer` the same way.

What per-user means:

- Each user of a Mac installs it for themselves; a new user on the Mac does not get it.
- Detection: look in the console user's home folder, for example
  `test -x "/Users/$user/Library/Application Support/open-charts/engine/open-charts"`, or
  `pkgutil --volume "/Users/$user" --pkg-info io.github.open-charts.addin`.
- Updates: install the newer `.pkg` over it, the same way.
- Uninstall: the user opens **Uninstall open-charts** in `~/Applications`, or run as the user:
  `"$HOME/Library/Application Support/open-charts/engine/open-charts" addin uninstall`. It removes the
  add-in, engine, pop-up app, bridge script, token, command, LaunchAgent and receipt. Settings and
  logs stay in `~/Library/Application Support/open-charts`; delete that folder for a full clean.

Until the package is signed and notarized, Gatekeeper warns when a user opens it by double-click.
A package your MDM copies to the Mac usually has no quarantine mark, so `installer` from a script
should not be stopped. Check this in a pilot.

### What the user is asked

1. **Macros.** PowerPoint may ask about macros in `open-charts.ppam`; the user clicks **Enable
   Macros**. PowerPoint for Mac's macro setting can be managed with a configuration profile for the
   `com.microsoft.Powerpoint` domain (key `VisualBasicMacroExecutionState`). A value that disables
   macros without warning blocks the add-in.
2. **Automation, once for the installer's registration.** If PowerPoint was not running during the
   install, the LaunchAgent registers the add-in through Apple events. macOS asks whether
   *open-charts* may control *Microsoft PowerPoint*. If the user says no, the agent gives up after
   10 tries; the user can then load the add-in in *Tools › PowerPoint Add-ins… › +*.
3. **Automation, once for Apply.** The first time the user applies data from the Edit data window,
   macOS asks: *"Microsoft PowerPoint" wants access to control "Microsoft PowerPoint"*. The engine
   runs as PowerPoint's child (started through PowerPoint's `AppleScriptTask`), so macOS treats
   PowerPoint as the one sending the Apple events. Until the user clicks **Allow**, Apply cannot
   update the slide.

The add-in needs **no other permission**. It does not need Accessibility, Input Monitoring, Screen
Recording or Full Disk Access. The Edit data window and its drag handles watch only mouse
events, and they read only window positions. The bridge script runs from PowerPoint's own
Application Scripts folder, which macOS lets PowerPoint use without a prompt.

### PPPC profile (Automation)

You can pre-approve prompt 3 with a Privacy Preferences Policy Control payload
(`com.apple.TCC.configuration-profile-policy`). It must come from a user-approved or supervised
MDM enrollment. The payload allows PowerPoint to send Apple events to itself:

```xml
<key>Services</key>
<dict>
  <key>AppleEvents</key>
  <array>
    <dict>
      <key>Identifier</key><string>com.microsoft.Powerpoint</string>
      <key>IdentifierType</key><string>bundleID</string>
      <key>CodeRequirement</key><string>PASTE: codesign -dr - "/Applications/Microsoft PowerPoint.app"</string>
      <key>AEReceiverIdentifier</key><string>com.microsoft.Powerpoint</string>
      <key>AEReceiverIdentifierType</key><string>bundleID</string>
      <key>AEReceiverCodeRequirement</key><string>PASTE: the same requirement</string>
      <key>Allowed</key><true/>
    </dict>
  </array>
</dict>
```

Take both code requirements from your own PowerPoint with
`codesign -dr - "/Applications/Microsoft PowerPoint.app"` (the text after `designated =>`). Jamf's
and other MDMs' PPPC editors fill these in from the app.

Prompt 2 comes from the engine executable, which is unsigned today. A PPPC entry needs a stable code
requirement, so prompt 2 cannot be pre-approved until the engine is signed with a Developer ID.
If you install while PowerPoint is running, the add-in is registered right away and the
LaunchAgent is not used.

## Office.js add-in: Microsoft 365 centralized deployment

The Office.js add-in installs nothing on the computer. Host its files on a https server and deploy
its `manifest.xml` from the Microsoft 365 admin center, under *Settings › Integrated apps*. The
steps, hosting needs and what the pane sends over the network are in
[OFFICEJS-HOSTING.md](OFFICEJS-HOSTING.md).

## Updates

- **Off by default.** No install checks for updates unless a user turns it on with
  `open-charts addin update on`.
- **Force it off.** Put a policy file with `{"updateCheck": false}` at:
  - Mac: `/Library/Application Support/open-charts/policy.json`
  - Windows: `%ProgramData%\open-charts\policy.json`

  With it, the add-in never checks. `open-charts addin update` and `open-charts addin update install`
  say "Updates are managed by your IT department".
- **What a user's update does.** `open-charts addin update install` downloads the platform's
  installer and `SHA256SUMS.txt` from the release. It checks the SHA-256 and the publisher
  signature (`pkgutil` and `spctl` on a Mac, `Get-AuthenticodeSignature` on Windows). Only then
  does it open the installer, which the user confirms. Until the packages are signed, it refuses:
  nothing unsigned is opened.
- **Managed machines:** deploy each new version with your tools, as above, and set the policy file.
